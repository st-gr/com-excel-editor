---
name: com-excel-editor
description: "Edit Excel workbooks (.xlsx, .xlsm) on Windows using COM automation with full Microsoft Excel capabilities including cell operations, formatting, formulas, charts, notes, threaded comments, and Notes-to-Comments conversion. Use this skill whenever the user needs to edit Excel files via COM, open or edit SharePoint or Teams-hosted Excel files via URL, convert legacy Excel notes into modern threaded comments, or perform Excel automation requiring the full COM API. Trigger for ANY Excel editing task on Windows — especially when SharePoint access, Notes-to-Comments conversion, formula preservation, or advanced formatting is involved. Also trigger when the user pastes a SharePoint or Teams URL ending in .xlsx or .xlsm."
---

# COM Excel Editor

Full-fidelity Excel workbook editing through Windows COM automation. Edits local files and SharePoint/Teams-hosted workbooks with formulas, formatting, notes, threaded comments, charts, and more.

## When to Use This Skill

**Use COM editing (this skill) when:**
- You need to **preserve formulas, macros, or VBA** in an existing workbook
- The workbook is hosted on **SharePoint or Teams** (opened by URL)
- You need to **convert legacy Notes into modern threaded Comments**
- You need to preserve **conditional formatting, charts, or pivot tables**
- The task requires full Excel fidelity (data validation, named ranges, etc.)

**Use openpyxl instead when:**
- Creating a new workbook from scratch (no existing file to modify)
- Running on Linux/Mac (no Excel installed)
- Only simple cell reads/writes are needed, no COM dependency desired

## Prerequisites

```bash
# Install pywin32 (provides win32com.client)
uv add pywin32
# or: pip install pywin32
```

- **Windows** with **Microsoft Excel** installed
- **SharePoint access**: user must be authenticated via Office 365 SSO or Windows credential manager

## Quick Start

The bundled script is at `<skill-directory>/scripts/com_excel_editor.py`. Copy it to your working directory or add the skill's `scripts/` folder to `sys.path`.

```python
import sys
sys.path.insert(0, "<skill-directory>/scripts")
from com_excel_editor import ExcelWorkbook

with ExcelWorkbook() as wb:
    wb.open("path/to/workbook.xlsx")
    sheet = wb.get_sheet("Sheet1")
    sheet.write_cell(2, 3, "Updated value")
    # auto-saves on exit, restores original settings
```

## Opening SharePoint / Teams Documents

Pass the sharing URL directly -- the script auto-parses it:

```python
url = "https://tenant.sharepoint.com/:x:/r/teams/Team/Shared%20Documents/General/report.xlsx?d=xxx&csf=1&web=1&e=xxx"

with ExcelWorkbook() as wb:
    wb.open(url)  # auto-converts to direct document path
    sheet = wb.get_sheet("Data")
    sheet.write_cell(1, 1, "Updated")
```

The parser extracts the path between `/:x:/r/` and `?`, URL-decodes it, and reconstructs a direct URL that Excel COM can open. The user must be authenticated to SharePoint (Office 365 SSO handles this automatically on domain-joined machines).

## Available Operations

### 1. Workbook Lifecycle

```python
wb = ExcelWorkbook(visible=False)   # create Excel instance
wb.open("file.xlsx")                # open local file
wb.open("https://sharepoint/...")   # open SharePoint file
wb.save()                           # save
wb.save_as("new.xlsx")             # save as
wb.export_pdf("output.pdf")        # export to PDF
wb.close()                          # close and quit Excel
```

### 2. Sheet Operations

```python
sheet = wb.get_sheet("Sheet1")      # get sheet by name
sheet = wb.get_sheet(1)             # get sheet by index (1-based)
names = wb.sheet_names              # list all sheet names
wb.add_sheet("NewSheet")           # add a new sheet
wb.delete_sheet("OldSheet")        # delete a sheet
```

### 3. Cell Operations

```python
value = sheet.read_cell(2, 3)            # read cell (row 2, col 3)
sheet.write_cell(2, 3, "Hello")          # write a string
sheet.write_cell(2, 3, 42)              # write a number
sheet.write_cell(2, 3, "=SUM(A1:A10)")  # write a formula

# Bulk read/write ranges
data = sheet.read_range(1, 1, 10, 5)    # read rows 1-10, cols 1-5
sheet.write_range(1, 1, [["A","B"],["C","D"]])  # write 2x2 block
```

### 4. Find & Replace

```python
sheet.find_replace("old", "new")                    # replace all in sheet
sheet.find_replace("old", "new", match_case=True)   # case-sensitive
```

### 5. Notes (Legacy Comments)

In modern Excel, "Notes" are the yellow sticky-note annotations (formerly called "Comments" pre-Office 365).

```python
note = sheet.get_note(2, 3)          # read note text (or None)
sheet.add_note(2, 3, "Review this")  # add/overwrite a note
sheet.delete_note(2, 3)             # delete a note
all_notes = sheet.get_all_notes()    # list of {row, col, text}
```

### 6. Threaded Comments (Modern)

Modern threaded comments support replies and @mentions. These replaced Notes as the primary annotation mechanism.

```python
sheet.add_comment(2, 3, "Please verify this figure")
comment = sheet.get_comment(2, 3)    # {author, text, replies}
sheet.reply_to_comment(2, 3, "Verified -- looks correct")
```

### 7. Notes-to-Comments Conversion

Convert legacy Notes into modern threaded Comments (see dedicated section below).

```python
sheet.convert_note_to_comment(2, 3)       # single cell
sheet.convert_all_notes_to_comments()     # entire sheet
wb.convert_all_notes_to_comments()        # entire workbook
```

### 8. Formatting

```python
sheet.format_cell(2, 3, bold=True, font_size=12, bg_color=(0, 0, 255))
sheet.autofit("A:D")                # autofit columns A through D
sheet.set_column_width("B", 25)     # set column B to width 25
```

### 9. Workbook Info

```python
title = wb.get_property("Title")
wb.set_property("Title", "Q4 Financial Report")
```

### 10. Raw COM Access

For anything not covered above, access the underlying COM objects directly:

```python
excel_app = wb.app   # COM Application object
excel_wb = wb.doc    # COM Workbook object

# Example: freeze panes on the active sheet
excel_wb.Sheets("Data").Activate()
excel_app.ActiveWindow.FreezePanes = True

# Example: insert a chart
chart = excel_wb.Sheets("Data").ChartObjects.Add(100, 50, 400, 300)
chart.Chart.SetSourceData(excel_wb.Sheets("Data").Range("A1:B10"))
```

## Notes-to-Comments Conversion

### Background

In modern Excel (Office 365 / Excel 2019+), Microsoft split the old "Comments" feature into two:

- **Notes** -- the legacy yellow sticky-note boxes. No threading, no @mentions. These are what older files contain.
- **Comments** -- modern threaded discussions anchored to cells. Support replies, @mentions, and resolve/reopen.

Many organizations need to migrate from Notes to Comments for collaboration. This is not possible through openpyxl or the Excel UI in bulk.

### How It Works

The conversion follows a 3-step process per cell:
1. **Read** the note text from the cell
2. **Delete** the note
3. **Add** a modern threaded comment with the same text

### Usage

```python
# Single cell
sheet.convert_note_to_comment(2, 3)

# All notes on one sheet
converted = sheet.convert_all_notes_to_comments()
print(f"Converted {converted} notes")

# All notes in the entire workbook
total = wb.convert_all_notes_to_comments()
print(f"Converted {total} notes across all sheets")
```

The original note author is preserved when available. If a note has no author metadata, the current user is used as the comment author.

## Critical Pitfalls

These are hard-won lessons from production debugging. The bundled script handles all of them internally, but if you write custom COM code, be aware:

### 1. Performance: Always Use the Context Manager

The context manager sets `ScreenUpdating = False` and `Calculation = xlCalculationManual` on entry, then restores both on exit. Without this, writing hundreds of cells triggers a screen repaint and full recalculation after every single write -- turning a 2-second operation into a 10-minute crawl.

### 2. Colors Are BGR, Not RGB

Excel COM uses BGR byte order for colors. To set a cell to red (`RGB 255, 0, 0`), you must pass `0x0000FF` (BGR). The bundled script's `format_cell` method accepts standard `(R, G, B)` tuples and converts internally.

### 3. 1-Based Indexing

Rows and columns in Excel COM start at 1, not 0. `Cells(1, 1)` is cell A1. The bundled script follows this convention -- `read_cell(1, 1)` reads A1.

### 4. No Tracked Changes in Excel

Unlike Word, Excel COM does not have a built-in tracked changes API. There is no equivalent to `TrackRevisions`. If you need an audit trail, implement it at the application level (e.g., log changes to a separate sheet).

## Excel Constants Reference

| Constant | Value | Use |
|---|---|---|
| `xlCalculationManual` | -4135 | Disable auto-recalculation |
| `xlCalculationAutomatic` | -4105 | Enable auto-recalculation |
| `xlUp` | -4162 | Direction: up |
| `xlDown` | -4121 | Direction: down |
| `xlToLeft` | -4159 | Direction: left |
| `xlToRight` | -4161 | Direction: right |
| `xlLastCell` | 11 | Special cell: last used cell |
| `xlWorkbookDefault` | 51 | Save as .xlsx |
| `xlOpenXMLWorkbookMacroEnabled` | 52 | Save as .xlsm |
| `xlTypePDF` | 0 | Export as PDF |

For the full Excel VBA object model: [Microsoft Excel VBA Reference](https://learn.microsoft.com/en-us/office/vba/api/overview/excel)
