# com-excel-editor

Full-fidelity Excel workbook editing through Windows COM automation -- formulas, formatting, notes, threaded comments, charts, and SharePoint integration.

## Why COM instead of openpyxl?

| Capability | COM (this skill) | openpyxl |
|---|---|---|
| Preserve formulas on save | Yes | Partial (loses volatile formulas) |
| Preserve macros / VBA | Yes | No (.xlsx only) |
| Open SharePoint / Teams files | Yes (by URL) | No |
| Notes-to-Comments conversion | Yes | No |
| Conditional formatting fidelity | Full | Partial |
| Charts and pivot tables | Full | Limited |
| Platform | Windows only | Cross-platform |

## Installation

### As a Claude Code skill

Place the `com-excel-editor/` folder under `~/.claude/skills/`. Claude Code will automatically discover it via the `SKILL.md` frontmatter.

### As a standalone module

```bash
pip install pywin32
# Copy scripts/com_excel_editor.py into your project
```

Requires Windows with Microsoft Excel installed.

## Quick Start

### Local file

```python
from com_excel_editor import ExcelWorkbook

with ExcelWorkbook() as wb:
    wb.open("report.xlsx")
    sheet = wb.get_sheet("Summary")
    sheet.write_cell(2, 3, "Updated")
```

### SharePoint / Teams URL

```python
from com_excel_editor import ExcelWorkbook

url = "https://tenant.sharepoint.com/:x:/r/teams/Team/Shared%20Documents/General/report.xlsx?d=xxx&csf=1&web=1&e=xxx"

with ExcelWorkbook() as wb:
    wb.open(url)
    sheet = wb.get_sheet("Data")
    value = sheet.read_cell(1, 1)
```

## Features

**Workbook lifecycle**: `open`, `save`, `save_as`, `export_pdf`, `close`

**Sheet operations**: `get_sheet`, `add_sheet`, `delete_sheet`, `sheet_names`

**Cell operations**: `read_cell`, `write_cell`, `read_range`, `write_range`

**Find & Replace**: `find_replace` with case sensitivity options

**Notes (legacy)**: `get_note`, `add_note`, `delete_note`, `get_all_notes`

**Threaded Comments (modern)**: `add_comment`, `get_comment`, `reply_to_comment`

**Notes-to-Comments conversion**: `convert_note_to_comment`, `convert_all_notes_to_comments`

**Formatting**: `format_cell`, `autofit`, `set_column_width`

**Workbook info**: `get_property`, `set_property`

**Raw COM access**: `wb.app` (Application), `wb.doc` (Workbook) for direct COM calls

## Notes-to-Comments Conversion

The key unique feature of this skill. Modern Excel has two annotation types:

- **Notes** -- legacy yellow sticky-notes (pre-Office 365 "Comments")
- **Comments** -- modern threaded discussions with replies and @mentions

This skill converts Notes to Comments in bulk, which is not possible through the Excel UI or openpyxl.

```python
with ExcelWorkbook() as wb:
    wb.open("legacy_workbook.xlsx")

    # Convert a single cell
    sheet = wb.get_sheet("Sheet1")
    sheet.convert_note_to_comment(2, 3)

    # Convert an entire sheet
    sheet.convert_all_notes_to_comments()

    # Convert the entire workbook
    total = wb.convert_all_notes_to_comments()
    print(f"Converted {total} notes across all sheets")
```

## Key Design Decisions

1. **Performance optimization** -- The context manager disables `ScreenUpdating` and sets `Calculation` to manual mode on entry. This prevents Excel from repainting and recalculating after every cell write, turning bulk operations from minutes to seconds.

2. **BGR color order** -- Excel COM uses BGR byte order internally. The `format_cell` method accepts standard `(R, G, B)` tuples and converts to BGR automatically, so you never have to think about it.

3. **1-based indexing** -- All row and column parameters are 1-based to match Excel's native COM convention. `read_cell(1, 1)` reads cell A1.

4. **No tracked changes** -- Unlike Word, Excel has no built-in tracked changes API. This is a platform limitation, not a missing feature. If you need an audit trail, log changes to a separate sheet.

## License

MIT
