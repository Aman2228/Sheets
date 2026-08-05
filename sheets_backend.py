"""
sheets_backend.py — makes a Google Sheet look like an openpyxl workbook.

Why: ocs_master.py's logic was written against openpyxl (ws.cell(row, col).value,
wb.save(), etc). Rather than rewrite every function's cell-access style, this
module loads the whole Google Sheet into memory as plain grids (same as
openpyxl does with a workbook), exposes the same .cell()/.max_row/.max_column/
.create_sheet() surface, and pushes everything back with ONE batched API call
per changed sheet on save(). This also avoids hammering the Sheets API rate
limit (60 req/min/user) — openpyxl-style code can call .cell() thousands of
times per run and it costs zero extra API calls here.
"""
import os
import json
import gspread
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def _client():
    """
    Auth via a service account. Credentials come from the
    GOOGLE_SERVICE_ACCOUNT_JSON env var (the full JSON key, as one line/string)
    — set this in your hosting provider's environment variables. Never commit
    the JSON file itself.
    """
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not raw:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON env var is not set.")
    info = json.loads(raw)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


class _Cell:
    __slots__ = ("_grid", "_r", "_c")

    def __init__(self, grid, r, c):
        self._grid, self._r, self._c = grid, r, c

    @property
    def value(self):
        row = self._grid.rows[self._r - 1] if self._r - 1 < len(self._grid.rows) else []
        return row[self._c - 1] if self._c - 1 < len(row) else None

    @value.setter
    def value(self, v):
        self._grid.set(self._r, self._c, v)


class Grid:
    """One worksheet, held in memory as a list-of-lists (1-indexed access)."""

    def __init__(self, title, values):
        self.title = title
        # values: list of rows (list of str). Normalize to a rectangular grid.
        self.rows = [list(r) for r in values]
        self.dirty = False

    @property
    def max_row(self):
        return len(self.rows)

    @property
    def max_column(self):
        return max((len(r) for r in self.rows), default=0)

    def cell(self, row, column):
        return _Cell(self, row, column)

    def set(self, r, c, v):
        while len(self.rows) < r:
            self.rows.append([])
        row = self.rows[r - 1]
        while len(row) < c:
            row.append(None)
        row[c - 1] = v
        self.dirty = True

    # openpyxl compat no-ops
    class _ColDims(dict):
        def __getitem__(self, k):
            return self.setdefault(k, type("W", (), {"width": None})())

    @property
    def column_dimensions(self):
        return Grid._ColDims()


class FakeWorkbook:
    """Mimics openpyxl.Workbook, backed by a real Google Sheet."""

    def __init__(self, spreadsheet_key):
        self._gc = _client()
        self._sh = self._gc.open_by_key(spreadsheet_key)
        self._grids = {}
        for ws in self._sh.worksheets():
            values = ws.get_all_values()
            self._grids[ws.title] = Grid(ws.title, values)

    @property
    def sheetnames(self):
        return list(self._grids.keys())

    def __getitem__(self, name):
        return self._grids[name]

    def __contains__(self, name):
        return name in self._grids

    def create_sheet(self, name):
        g = Grid(name, [])
        self._grids[name] = g
        g.dirty = True  # force creation on save
        return g

    def save(self, _unused_path=None):
        """Push every changed grid back to Google Sheets in one call each."""
        existing_titles = {ws.title: ws for ws in self._sh.worksheets()}
        for name, grid in self._grids.items():
            if not grid.dirty:
                continue
            if name not in existing_titles:
                rows_needed = max(len(grid.rows), 1)
                cols_needed = max(grid.max_column, 1)
                ws = self._sh.add_worksheet(title=name, rows=rows_needed + 10, cols=cols_needed + 5)
            else:
                ws = existing_titles[name]
                # grow the sheet if we added rows/cols beyond current size
                need_rows, need_cols = len(grid.rows), grid.max_column
                if ws.row_count < need_rows or ws.col_count < need_cols:
                    ws.resize(rows=max(ws.row_count, need_rows) + 5,
                              cols=max(ws.col_count, need_cols) + 2)
            if grid.rows:
                # pad rows to rectangular shape gspread expects
                width = grid.max_column
                padded = [r + [None] * (width - len(r)) for r in grid.rows]
                ws.update(padded, "A1")
            grid.dirty = False


def load_workbook(spreadsheet_key):
    return FakeWorkbook(spreadsheet_key)
