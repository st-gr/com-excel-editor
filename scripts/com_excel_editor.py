"""
com_excel_editor.py - Excel COM Automation via pywin32

A production-ready module for automating Microsoft Excel through COM.
Provides workbook lifecycle management, sheet operations, cell reading/writing,
find/replace, formatting, notes (legacy comments), threaded comments (modern),
notes-to-comments conversion, and PDF export.

Critical design notes (from real production debugging):
  - Excel uses 1-based row/column indexing throughout.
  - Colours in Excel COM are BGR (Blue-Green-Red), not RGB. Use _rgb_to_bgr().
  - Performance optimization (ScreenUpdating, Calculation, etc.) is saved on
    entry and restored on exit to avoid corrupting the user's settings.
  - SharePoint URLs use a letter indicator (e.g. :x: for Excel, :w: for Word)
    in the path; the parser handles any single-letter indicator.
  - The legacy "Comment" object in Excel COM is what users see as yellow
    sticky "Notes". Modern threaded comments use the CommentThreaded API.
  - Context manager auto-saves on clean exit and always restores perf settings.

Requires: pywin32 (pip install pywin32), Microsoft Excel on Windows.
"""

import os
import re
from urllib.parse import unquote, urlparse

import pywintypes
import win32com.client

__all__ = [
    "parse_sharepoint_url",
    "SheetWrapper",
    "ExcelWorkbook",
    # Constants (commonly needed by callers)
    "xlWorkbookDefault",
    "xlCSV",
    "xlCSVUTF8",
    "xlOpenXMLWorkbookMacroEnabled",
    "xlTypePDF",
    "xlCalculationAutomatic",
    "xlCalculationManual",
]

# ---------------------------------------------------------------------------
# Excel Constants
# ---------------------------------------------------------------------------

# Direction (used with End() method for navigation)
xlUp = -4162
xlDown = -4121
xlToLeft = -4159
xlToRight = -4161

# Match type for Find
xlWhole = 1
xlPart = 2

# Search order
xlByRows = 1
xlByColumns = 2

# Calculation mode
xlCalculationAutomatic = -4105
xlCalculationManual = -4135
xlCalculationSemiautomatic = 2

# Save formats
xlWorkbookDefault = 51          # .xlsx
xlCSV = 6                       # .csv (ANSI)
xlCSVUTF8 = 62                  # .csv (UTF-8)
xlOpenXMLWorkbookMacroEnabled = 52  # .xlsm

# PDF export
xlTypePDF = 0

# Cell types (for SpecialCells)
xlCellTypeConstants = 2
xlCellTypeFormulas = -4123
xlCellTypeLastCell = 11

# Horizontal alignment
xlHAlignLeft = -4131
xlHAlignCenter = -4108
xlHAlignRight = -4152

# Borders
xlEdgeBottom = 9
xlEdgeLeft = 7
xlEdgeRight = 10
xlEdgeTop = 8
xlContinuous = 1
xlNone = -4142

# Colours (BGR! -- Excel stores colours as 0xBBGGRR, not 0xRRGGBB)
xlColorBlack = 0x000000
xlColorRed = 0x0000FF
xlColorBlue = 0xFF0000
xlColorGreen = 0x00FF00
xlColorWhite = 0xFFFFFF


# ---------------------------------------------------------------------------
# Utility: SharePoint / Teams URL parsing
# ---------------------------------------------------------------------------

def parse_sharepoint_url(url: str) -> str:
    """Convert a SharePoint/Teams sharing URL to a direct document path.

    Handles Excel's ``:x:`` indicator as well as any other single-letter
    indicator (``:w:``, ``:p:``, etc.).

    Input example::

        https://tenant.sharepoint.com/:x:/r/teams/Team/Shared%20Documents/General/data.xlsx?d=xxx&csf=1&web=1&e=xxx

    Output::

        https://tenant.sharepoint.com/teams/Team/Shared Documents/General/data.xlsx
    """
    parsed = urlparse(url)

    # Pattern: /:<letter>:/r/<path>  (e.g. :x: for Excel, :w: for Word)
    match = re.search(r"/:[a-z]:/r/(.+?)(?:\?|$)", parsed.path)
    if match:
        doc_path = unquote(match.group(1))
        return f"{parsed.scheme}://{parsed.netloc}/{doc_path}"

    # Already a direct URL -- strip query params and decode
    clean = unquote(parsed.path)
    return f"{parsed.scheme}://{parsed.netloc}{clean}"


def _rgb_to_bgr(r: int, g: int, b: int) -> int:
    """Convert an (R, G, B) tuple to Excel's BGR long colour value.

    Excel stores interior and font colours as a single long integer in
    BGR byte order: ``0xBBGGRR``.

    Example::

        >>> _rgb_to_bgr(255, 0, 0)   # red
        255
        >>> _rgb_to_bgr(0, 0, 255)   # blue
        16711680
    """
    return (b << 16) | (g << 8) | r


# ---------------------------------------------------------------------------
# SheetWrapper
# ---------------------------------------------------------------------------

class SheetWrapper:
    """Convenience wrapper around an Excel COM Worksheet object.

    All row and column indices are **1-based**, matching Excel's native
    addressing.  Range addresses use standard Excel notation (e.g. ``"A1:C10"``).
    """

    def __init__(self, sheet, workbook: "ExcelWorkbook"):
        self._sheet = sheet
        self._workbook = workbook

    # -- properties --------------------------------------------------------

    @property
    def name(self) -> str:
        """The worksheet tab name."""
        return self._sheet.Name

    @property
    def com_object(self):
        """The raw COM Worksheet object for advanced use."""
        return self._sheet

    # -- cell operations (1-based row/col) ---------------------------------

    def read_cell(self, row: int, col: int):
        """Read the value of a single cell.

        Returns the cell's ``.Value`` which may be a string, number,
        datetime, boolean, or None for empty cells.
        """
        return self._sheet.Cells(row, col).Value

    def write_cell(self, row: int, col: int, value):
        """Write a value to a single cell.

        Accepts strings, numbers, dates, booleans, or formulas (strings
        starting with ``=``).
        """
        self._sheet.Cells(row, col).Value = value

    def read_range(self, address: str) -> list[list]:
        """Read a rectangular range and return as a list of lists.

        Each inner list represents one row.  Empty cells appear as None.

        Example::

            data = sheet.read_range("A1:C3")
            # [[1, 'Name', 100], [2, 'Other', 200], [3, None, 300]]
        """
        rng = self._sheet.Range(address)
        raw = rng.Value
        if raw is None:
            return [[]]
        # Single cell returns a scalar, not a tuple
        if not isinstance(raw, tuple):
            return [[raw]]
        # COM returns a tuple of tuples (rows of columns)
        return [list(row) for row in raw]

    def write_range(self, address: str, data: list[list]):
        """Write a 2D list of lists into a rectangular range.

        The dimensions of *data* must match the range specified by *address*.
        Each inner list is one row.

        Example::

            sheet.write_range("A1:B2", [[1, 2], [3, 4]])
        """
        rng = self._sheet.Range(address)
        rng.Value = data

    # -- dimensions --------------------------------------------------------

    def used_range(self) -> tuple[int, int, int, int]:
        """Return the bounding box of the used range.

        Returns ``(first_row, first_col, last_row, last_col)`` as 1-based
        integers.  If the sheet is completely empty, returns ``(1, 1, 1, 1)``.
        """
        ur = self._sheet.UsedRange
        first_row = ur.Row
        first_col = ur.Column
        last_row = first_row + ur.Rows.Count - 1
        last_col = first_col + ur.Columns.Count - 1
        return (first_row, first_col, last_row, last_col)

    def last_row(self, col: int = 1) -> int:
        """Return the last non-empty row number in the given column.

        Uses the standard ``End(xlUp)`` technique: starts from the very
        last row in the column and searches upward.  Returns 0 if the
        entire column is empty.
        """
        max_row = self._sheet.Rows.Count  # 1048576 for modern Excel
        cell = self._sheet.Cells(max_row, col).End(xlUp)
        # If the result is row 1, check if that cell is actually empty
        if cell.Row == 1 and cell.Value is None:
            return 0
        return cell.Row

    def last_col(self, row: int = 1) -> int:
        """Return the last non-empty column number in the given row.

        Uses ``End(xlToLeft)`` from the rightmost column.  Returns 0 if the
        entire row is empty.
        """
        max_col = self._sheet.Columns.Count  # 16384 for modern Excel
        cell = self._sheet.Cells(row, max_col).End(xlToLeft)
        if cell.Column == 1 and cell.Value is None:
            return 0
        return cell.Column

    # -- notes (legacy "comments" -- yellow sticky notes) ------------------

    def get_note(self, row: int, col: int) -> str | None:
        """Read the legacy note (yellow sticky) on a cell, or None."""
        cell = self._sheet.Cells(row, col)
        try:
            comment = cell.Comment
            if comment is None:
                return None
            return comment.Text()
        except pywintypes.com_error:
            return None

    def _ensure_calc_auto(self):
        """Temporarily set Calculation to Automatic if it's Manual.

        Excel COM's AddComment/AddCommentThreaded fail with 'RPC call
        failed' when Calculation is set to Manual.  This method saves the
        current mode and restores it after the comment operation via
        ``_restore_calc()``.
        """
        app = self._workbook.app
        try:
            self._saved_calc = app.Calculation
            if self._saved_calc == xlCalculationManual:
                app.Calculation = xlCalculationAutomatic
        except pywintypes.com_error:
            self._saved_calc = None

    def _restore_calc(self):
        """Restore the Calculation mode saved by ``_ensure_calc_auto()``."""
        if getattr(self, "_saved_calc", None) is not None:
            try:
                self._workbook.app.Calculation = self._saved_calc
            except pywintypes.com_error:
                pass
            self._saved_calc = None

    def add_note(self, row: int, col: int, text: str):
        """Add a legacy note (yellow sticky) to a cell.

        If a note already exists it is replaced.

        Uses ``Range.NoteText()`` instead of ``Range.AddComment()``
        because AddComment triggers RPC failures on machines with
        corporate COM add-ins (e.g. Enterprise Connect).
        """
        cell = self._sheet.Cells(row, col)
        try:
            if cell.Comment is not None:
                cell.Comment.Delete()
        except pywintypes.com_error:
            pass
        cell.NoteText(text)
        print(f"  Note added at ({row},{col})")

    def delete_note(self, row: int, col: int):
        """Delete the legacy note on a cell, if present."""
        cell = self._sheet.Cells(row, col)
        try:
            if cell.Comment is not None:
                cell.Comment.Delete()
                print(f"  Note deleted at ({row},{col})")
        except pywintypes.com_error:
            pass

    def get_all_notes(self) -> list[dict]:
        """Return all legacy notes on this sheet.

        Returns a list of dicts, each with keys ``row``, ``col``, ``text``.
        """
        results = []
        try:
            comments = self._sheet.Comments
            if comments is None:
                return results
            for i in range(1, comments.Count + 1):
                comment = comments(i)
                cell = comment.Parent
                results.append({
                    "row": cell.Row,
                    "col": cell.Column,
                    "text": comment.Text(),
                })
        except pywintypes.com_error:
            pass
        return results

    # -- threaded comments (modern collaborative comments) -----------------

    def get_comment(self, row: int, col: int) -> str | None:
        """Read the root text of the modern threaded comment on a cell.

        Returns None if no threaded comment exists.
        """
        cell = self._sheet.Cells(row, col)
        try:
            ct = cell.CommentThreaded
            if ct is None:
                return None
            return ct.Text()
        except (pywintypes.com_error, AttributeError):
            # AttributeError: CommentThreaded not available on older Excel
            return None

    def add_comment(self, row: int, col: int, text: str):
        """Add a modern threaded comment to a cell.

        If a threaded comment already exists it is deleted first.
        """
        cell = self._sheet.Cells(row, col)
        try:
            if cell.CommentThreaded is not None:
                cell.CommentThreaded.Delete()
        except (pywintypes.com_error, AttributeError):
            pass
        self._ensure_calc_auto()
        try:
            cell.AddCommentThreaded(text)
            print(f"  Threaded comment added at ({row},{col})")
        except (pywintypes.com_error, AttributeError) as exc:
            print(f"  [WARN] Could not add threaded comment at ({row},{col}): {exc}")
        finally:
            self._restore_calc()

    def reply_to_comment(self, row: int, col: int, text: str):
        """Add a reply to the existing threaded comment on a cell.

        Does nothing if no threaded comment exists.
        """
        cell = self._sheet.Cells(row, col)
        try:
            ct = cell.CommentThreaded
            if ct is None:
                print(f"  [WARN] No threaded comment at ({row},{col}) to reply to")
                return
            ct.AddReply(text)
            print(f"  Reply added to comment at ({row},{col})")
        except (pywintypes.com_error, AttributeError) as exc:
            print(f"  [WARN] Could not reply to comment at ({row},{col}): {exc}")

    def delete_comment(self, row: int, col: int):
        """Delete the modern threaded comment on a cell, if present."""
        cell = self._sheet.Cells(row, col)
        try:
            ct = cell.CommentThreaded
            if ct is not None:
                ct.Delete()
                print(f"  Threaded comment deleted at ({row},{col})")
        except (pywintypes.com_error, AttributeError):
            pass

    # -- notes-to-comments conversion --------------------------------------

    def convert_note_to_comment(self, row: int, col: int) -> bool:
        """Convert a legacy note to a modern threaded comment on one cell.

        Reads the note text, deletes the note, and creates a threaded
        comment with the same text.  Returns True on success, False if
        no note exists or conversion fails.
        """
        cell = self._sheet.Cells(row, col)
        try:
            if cell.Comment is None:
                return False
            note_text = cell.Comment.Text()
            cell.Comment.Delete()
            self._ensure_calc_auto()
            try:
                cell.AddCommentThreaded(note_text)
            finally:
                self._restore_calc()
            print(f"  Converted note to comment at ({row},{col})")
            return True
        except (pywintypes.com_error, AttributeError) as exc:
            print(f"  [WARN] convert_note_to_comment({row},{col}) failed: {exc}")
            return False

    def convert_all_notes_to_comments(self) -> int:
        """Convert every legacy note on this sheet to a modern threaded comment.

        Iterates the sheet's Comments collection (legacy notes), reads each
        note's text and parent cell, deletes the note, and creates a threaded
        comment.  Returns the number of notes successfully converted.
        """
        converted = 0
        try:
            comments = self._sheet.Comments
            if comments is None or comments.Count == 0:
                return 0
        except pywintypes.com_error:
            return 0

        # Collect note data first, then modify (avoids mutating during iteration)
        notes_data = []
        for i in range(1, comments.Count + 1):
            try:
                comment = comments(i)
                cell = comment.Parent
                notes_data.append({
                    "row": cell.Row,
                    "col": cell.Column,
                    "text": comment.Text(),
                })
            except pywintypes.com_error:
                continue

        for note in notes_data:
            row, col, text = note["row"], note["col"], note["text"]
            try:
                cell = self._sheet.Cells(row, col)
                # Delete the legacy note
                if cell.Comment is not None:
                    cell.Comment.Delete()
                # Add modern threaded comment
                cell.AddCommentThreaded(text)
                converted += 1
                print(f"  Converted note to comment at ({row},{col})")
            except (pywintypes.com_error, AttributeError) as exc:
                print(f"  [WARN] Failed to convert note at ({row},{col}): {exc}")

        return converted

    # -- formatting --------------------------------------------------------

    def format_cell(self, row: int, col: int, bold: bool = None,
                    italic: bool = None, font_name: str = None,
                    font_size: float = None, color_rgb: tuple = None,
                    bg_color_rgb: tuple = None, number_format: str = None):
        """Apply formatting to a single cell.

        Parameters:
            bold: Set font bold.
            italic: Set font italic.
            font_name: Font family name (e.g. ``"Calibri"``).
            font_size: Font size in points.
            color_rgb: Font colour as ``(R, G, B)`` tuple (converted to BGR).
            bg_color_rgb: Cell background colour as ``(R, G, B)`` tuple.
            number_format: Excel number format string (e.g. ``"#,##0.00"``).
        """
        cell = self._sheet.Cells(row, col)
        if bold is not None:
            cell.Font.Bold = bold
        if italic is not None:
            cell.Font.Italic = italic
        if font_name is not None:
            cell.Font.Name = font_name
        if font_size is not None:
            cell.Font.Size = font_size
        if color_rgb is not None:
            cell.Font.Color = _rgb_to_bgr(*color_rgb)
        if bg_color_rgb is not None:
            cell.Interior.Color = _rgb_to_bgr(*bg_color_rgb)
        if number_format is not None:
            cell.NumberFormat = number_format

    def set_column_width(self, col: int, width: float):
        """Set the width of a column (1-based index).

        Width is in Excel's character-width units (approximately the width
        of the ``0`` character in the default font).
        """
        self._sheet.Columns(col).ColumnWidth = width

    def set_row_height(self, row: int, height: float):
        """Set the height of a row in points."""
        self._sheet.Rows(row).RowHeight = height

    def autofit_columns(self, start_col: int = None, end_col: int = None):
        """Auto-fit column widths to their content.

        If *start_col* and *end_col* are both None, auto-fits the entire
        used range.  Otherwise, auto-fits the specified column span.
        """
        if start_col is None and end_col is None:
            self._sheet.UsedRange.Columns.AutoFit()
        else:
            sc = start_col or 1
            ec = end_col or sc
            self._sheet.Range(
                self._sheet.Cells(1, sc),
                self._sheet.Cells(1, ec),
            ).EntireColumn.AutoFit()
        print(f"  Auto-fit columns on '{self.name}'")


# ---------------------------------------------------------------------------
# ExcelWorkbook
# ---------------------------------------------------------------------------

class ExcelWorkbook:
    """High-level Excel COM automation wrapper.

    Supports local files and SharePoint/Teams URLs.  Use as a context
    manager for automatic performance optimization and cleanup::

        with ExcelWorkbook(r"C:\\data.xlsx") as wb:
            sheet = wb.get_sheet(1)
            print(sheet.read_cell(1, 1))
            sheet.write_cell(1, 1, "Updated")
            # auto-saves on clean exit
    """

    def __init__(self, path_or_url: str = None, visible: bool = False):
        self._app = win32com.client.Dispatch("Excel.Application")
        self._app.Visible = visible
        self._wb = None

        # Original performance settings (populated by _optimize_start)
        self._orig_screen_updating = None
        self._orig_calculation = None
        self._orig_display_alerts = None
        self._orig_enable_events = None

        if path_or_url:
            self.open(path_or_url)

    # -- context manager ---------------------------------------------------

    def __enter__(self):
        self._optimize_start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            if exc_type is None and self._wb is not None:
                self.save()
        finally:
            self._optimize_end()
            self.close(save=False)
        return False  # do not suppress exceptions

    # -- performance optimization ------------------------------------------

    def _optimize_start(self):
        """Disable screen updating, auto-calculation, alerts, and events.

        Each setting is applied independently so a failure in one (e.g.
        Calculation requires an open workbook) does not block the others.
        """
        for attr, value, store in [
            ("ScreenUpdating", False, "_orig_screen_updating"),
            ("DisplayAlerts", False, "_orig_display_alerts"),
            ("EnableEvents", False, "_orig_enable_events"),
        ]:
            try:
                setattr(self, store, getattr(self._app, attr))
                setattr(self._app, attr, value)
            except pywintypes.com_error:
                pass
        # Calculation requires at least one open workbook
        try:
            self._orig_calculation = self._app.Calculation
            self._app.Calculation = xlCalculationManual
        except pywintypes.com_error:
            pass  # no workbook open yet — will be set after open()

    def _optimize_end(self):
        """Restore the original performance settings saved by ``_optimize_start()``.

        Always called in a finally block -- silently ignores errors so
        that cleanup never masks the original exception.
        """
        try:
            if self._orig_screen_updating is not None:
                self._app.ScreenUpdating = self._orig_screen_updating
            if self._orig_calculation is not None:
                self._app.Calculation = self._orig_calculation
            if self._orig_display_alerts is not None:
                self._app.DisplayAlerts = self._orig_display_alerts
            if self._orig_enable_events is not None:
                self._app.EnableEvents = self._orig_enable_events
        except pywintypes.com_error:
            pass

    # -- workbook lifecycle ------------------------------------------------

    def open(self, path_or_url: str) -> "ExcelWorkbook":
        """Open a local file or a SharePoint/Teams URL.

        SharePoint sharing URLs (containing ``/:x:/r/`` or similar) are
        automatically parsed into direct document paths before being
        passed to COM.
        """
        if "://" in path_or_url:
            target = parse_sharepoint_url(path_or_url)
        else:
            target = os.path.abspath(path_or_url)

        print(f"Opening: {target}")
        try:
            self._wb = self._app.Workbooks.Open(target)
        except pywintypes.com_error as exc:
            print(f"[ERROR] Could not open '{target}': {exc}")
            raise
        return self

    def new(self) -> "ExcelWorkbook":
        """Create a new blank workbook."""
        self._wb = self._app.Workbooks.Add()
        print("Created new workbook")
        return self

    def save(self):
        """Save the active workbook."""
        if self._wb is None:
            return
        print("Saving workbook...")
        try:
            self._wb.Save()
        except pywintypes.com_error as exc:
            print(f"[ERROR] Save failed: {exc}")
            raise

    def save_as(self, path: str, file_format: int = None):
        """Save the workbook to a new path.

        *file_format* accepts Excel constants such as ``xlWorkbookDefault``
        (51 for .xlsx) or ``xlCSVUTF8`` (62).  When ``None`` the format is
        inferred from the extension by Excel.
        """
        if self._wb is None:
            return
        abs_path = os.path.abspath(path)
        print(f"Saving as: {abs_path}")
        kwargs = {"Filename": abs_path}
        if file_format is not None:
            kwargs["FileFormat"] = file_format
        try:
            self._wb.SaveAs(**kwargs)
        except pywintypes.com_error as exc:
            print(f"[ERROR] SaveAs failed: {exc}")
            raise

    def export_pdf(self, path: str):
        """Export the entire workbook as a PDF file.

        Uses ``ExportAsFixedFormat`` with ``xlTypePDF``.
        """
        if self._wb is None:
            return
        abs_path = os.path.abspath(path)
        print(f"Exporting PDF: {abs_path}")
        try:
            self._wb.ExportAsFixedFormat(
                Type=xlTypePDF,
                Filename=abs_path,
                Quality=0,  # xlQualityStandard
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
        except pywintypes.com_error as exc:
            print(f"[ERROR] PDF export failed: {exc}")
            raise

    def close(self, save: bool = True):
        """Close the workbook and quit the Excel application.

        If *save* is True the workbook is saved before closing.
        """
        try:
            if self._wb is not None:
                self._wb.Close(SaveChanges=save)
                self._wb = None
        except pywintypes.com_error:
            pass
        try:
            self._app.Quit()
        except pywintypes.com_error:
            pass

    # -- raw COM access ----------------------------------------------------

    @property
    def doc(self):
        """The raw COM Workbook object."""
        return self._wb

    @property
    def app(self):
        """The raw COM Application object."""
        return self._app

    # -- sheet management --------------------------------------------------

    @property
    def sheet_count(self) -> int:
        """Number of worksheets in the workbook."""
        return self._wb.Worksheets.Count

    @property
    def sheet_names(self) -> list[str]:
        """List of worksheet names in order."""
        return [self._wb.Worksheets(i).Name
                for i in range(1, self._wb.Worksheets.Count + 1)]

    def get_sheet(self, index_or_name) -> SheetWrapper:
        """Return a ``SheetWrapper`` for the given sheet.

        *index_or_name* can be a 1-based integer index or a string sheet
        name.
        """
        try:
            sheet = self._wb.Worksheets(index_or_name)
            return SheetWrapper(sheet, self)
        except pywintypes.com_error as exc:
            print(f"[ERROR] get_sheet({index_or_name!r}): {exc}")
            raise

    def add_sheet(self, name: str = None) -> SheetWrapper:
        """Add a new worksheet to the workbook.

        The sheet is added after the last existing sheet.  If *name* is
        provided, the sheet is renamed.

        Returns a ``SheetWrapper`` for the new sheet.
        """
        after = self._wb.Worksheets(self._wb.Worksheets.Count)
        new_sheet = self._wb.Worksheets.Add(After=after)
        if name is not None:
            new_sheet.Name = name
        print(f"  Added sheet: '{new_sheet.Name}'")
        return SheetWrapper(new_sheet, self)

    def delete_sheet(self, index_or_name):
        """Delete a worksheet by index (1-based) or name.

        DisplayAlerts should be False (set by _optimize_start) to suppress
        the confirmation dialog.
        """
        try:
            sheet = self._wb.Worksheets(index_or_name)
            sheet_name = sheet.Name
            # Ensure no confirmation prompt
            orig_alerts = self._app.DisplayAlerts
            self._app.DisplayAlerts = False
            sheet.Delete()
            self._app.DisplayAlerts = orig_alerts
            print(f"  Deleted sheet: '{sheet_name}'")
        except pywintypes.com_error as exc:
            print(f"[ERROR] delete_sheet({index_or_name!r}): {exc}")
            raise

    @property
    def active_sheet(self) -> SheetWrapper:
        """Return a ``SheetWrapper`` for the currently active sheet."""
        return SheetWrapper(self._app.ActiveSheet, self)

    # -- find & replace ----------------------------------------------------

    def find_replace(self, find_text: str, replace_text: str,
                     sheet=None, match_case: bool = False) -> int:
        """Find and replace text across the workbook or a specific sheet.

        If *sheet* is None, iterates all sheets.  *sheet* can be a
        ``SheetWrapper``, a sheet name, or a 1-based index.

        Returns the total number of replacements made (estimated by
        counting matches before replacing, since Excel's Replace method
        does not return a count).
        """
        sheets_to_search = []
        if sheet is not None:
            if isinstance(sheet, SheetWrapper):
                sheets_to_search.append(sheet.com_object)
            else:
                sheets_to_search.append(self._wb.Worksheets(sheet))
        else:
            for i in range(1, self._wb.Worksheets.Count + 1):
                sheets_to_search.append(self._wb.Worksheets(i))

        total = 0
        for ws in sheets_to_search:
            # Count existing matches first
            count = 0
            try:
                first_found = ws.Cells.Find(
                    What=find_text,
                    LookIn=-4163,  # xlValues
                    LookAt=xlPart,
                    MatchCase=match_case,
                )
                if first_found is not None:
                    count = 1
                    first_addr = first_found.Address
                    nxt = ws.Cells.FindNext(first_found)
                    while nxt is not None and nxt.Address != first_addr:
                        count += 1
                        nxt = ws.Cells.FindNext(nxt)
            except pywintypes.com_error:
                count = 0

            if count > 0:
                try:
                    ws.Cells.Replace(
                        What=find_text,
                        Replacement=replace_text,
                        LookAt=xlPart,
                        MatchCase=match_case,
                    )
                    total += count
                    print(f"  Replace on '{ws.Name}': '{find_text}' -> "
                          f"'{replace_text}' ({count} matches)")
                except pywintypes.com_error as exc:
                    print(f"  [ERROR] Replace failed on '{ws.Name}': {exc}")

        if total == 0:
            print(f"  Find/Replace: '{find_text}' not found")
        return total

    def find_value(self, value, sheet=None) -> tuple | None:
        """Find the first cell containing *value*.

        Searches across all sheets unless *sheet* is specified.

        Returns ``(row, col, sheet_name)`` or None if not found.
        """
        sheets_to_search = []
        if sheet is not None:
            if isinstance(sheet, SheetWrapper):
                sheets_to_search.append(sheet.com_object)
            else:
                sheets_to_search.append(self._wb.Worksheets(sheet))
        else:
            for i in range(1, self._wb.Worksheets.Count + 1):
                sheets_to_search.append(self._wb.Worksheets(i))

        for ws in sheets_to_search:
            try:
                found = ws.Cells.Find(
                    What=value,
                    LookIn=-4163,  # xlValues
                    LookAt=xlPart,
                )
                if found is not None:
                    return (found.Row, found.Column, ws.Name)
            except pywintypes.com_error:
                continue
        return None

    # -- notes-to-comments (all sheets) ------------------------------------

    def convert_all_notes_to_comments(self) -> int:
        """Convert every legacy note in the entire workbook to a modern
        threaded comment.

        Iterates all sheets and delegates to each ``SheetWrapper``'s
        ``convert_all_notes_to_comments()`` method.

        Returns the total number of notes converted.
        """
        total = 0
        for i in range(1, self._wb.Worksheets.Count + 1):
            ws = SheetWrapper(self._wb.Worksheets(i), self)
            count = ws.convert_all_notes_to_comments()
            if count > 0:
                print(f"  Sheet '{ws.name}': converted {count} notes")
            total += count
        if total == 0:
            print("  No legacy notes found in workbook")
        else:
            print(f"  Total notes converted: {total}")
        return total

    # -- workbook properties -----------------------------------------------

    def get_property(self, name: str):
        """Read a built-in document property (e.g. ``'Title'``, ``'Author'``).

        Returns None if the property does not exist or cannot be read.
        """
        try:
            return self._wb.BuiltinDocumentProperties(name).Value
        except pywintypes.com_error:
            return None

    def set_property(self, name: str, value):
        """Set a built-in document property.

        Common property names: ``Title``, ``Subject``, ``Author``,
        ``Keywords``, ``Comments``, ``Category``, ``Manager``, ``Company``.
        """
        try:
            self._wb.BuiltinDocumentProperties(name).Value = value
            print(f"  Property '{name}' set to '{value}'")
        except pywintypes.com_error as exc:
            print(f"  [ERROR] set_property('{name}'): {exc}")
