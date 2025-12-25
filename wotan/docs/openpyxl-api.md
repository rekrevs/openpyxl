# openpyxl API Overview

A comprehensive reference to the openpyxl public API.

## 1. Main Entry Points

### Core Functions

```python
from openpyxl import Workbook, load_workbook
```

| Function | Description |
|----------|-------------|
| `Workbook(write_only=False, iso_dates=False)` | Create a new workbook |
| `load_workbook(filename, read_only=False, keep_vba=False, data_only=False, keep_links=True, rich_text=False)` | Load an existing .xlsx file |

### Load Options

| Parameter | Description |
|-----------|-------------|
| `read_only` | Optimized for reading large files (streaming) |
| `data_only` | Return formula results instead of formulas |
| `keep_vba` | Preserve VBA macros (.xlsm files) |
| `keep_links` | Preserve external links |
| `rich_text` | Preserve rich text formatting in cells |

---

## 2. Workbook (`openpyxl.workbook.Workbook`)

### Key Properties

| Property | Type | Description |
|----------|------|-------------|
| `active` | Worksheet | Get/set active worksheet |
| `worksheets` | list | All worksheets |
| `sheetnames` | list | Sheet name strings |
| `chartsheets` | list | All chartsheets |
| `defined_names` | DefinedNameDict | Named ranges/formulas |
| `properties` | DocumentProperties | Title, creator, dates, etc. |
| `security` | DocumentSecurity | Document protection |
| `read_only` | bool | Read-only mode flag |
| `write_only` | bool | Write-only mode flag |
| `data_only` | bool | Data-only mode flag |
| `epoch` | datetime | Date system epoch |
| `encoding` | str | Character encoding |
| `vba_archive` | ZipFile | VBA content (if preserved) |
| `loaded_theme` | Element | Theme colors |
| `calculation` | CalcProperties | Calculation settings |

### Key Methods

| Method | Description |
|--------|-------------|
| `create_sheet(title=None, index=None)` | Add a new worksheet |
| `remove(worksheet)` | Remove a worksheet |
| `move_sheet(sheet, offset=0)` | Move sheet position |
| `copy_worksheet(from_worksheet)` | Duplicate a worksheet |
| `create_chartsheet(title=None, index=None)` | Add a chartsheet |
| `add_named_style(style)` | Add a reusable named style |
| `save(filename)` | Save to file |
| `close()` | Close workbook |

### Access Patterns

```python
wb = Workbook()
ws = wb.active                    # Active sheet
ws = wb['Sheet1']                 # By name
ws = wb.worksheets[0]             # By index
'Sheet1' in wb                    # Check existence
for ws in wb:                     # Iterate sheets
    print(ws.title)
```

---

## 3. Worksheet (`openpyxl.worksheet.worksheet.Worksheet`)

### Key Properties

| Property | Type | Description |
|----------|------|-------------|
| `title` | str | Sheet name |
| `parent` | Workbook | Parent workbook |
| `dimensions` | str | Data range (e.g., "A1:Z100") |
| `min_row`, `max_row` | int | Row boundaries |
| `min_column`, `max_column` | int | Column boundaries |
| `merged_cells` | MergedCellRange | Merged cell ranges |
| `row_dimensions` | dict | Row heights/formatting |
| `column_dimensions` | dict | Column widths/formatting |
| `sheet_state` | str | 'visible', 'hidden', 'veryHidden' |
| `freeze_panes` | str | Frozen pane cell reference |
| `auto_filter` | AutoFilter | Filter settings |
| `data_validations` | DataValidationList | Validation rules |
| `conditional_formatting` | ConditionalFormattingList | CF rules |
| `tables` | TableList | Excel tables |
| `page_setup` | PageSetup | Print settings |
| `page_margins` | PageMargins | Print margins |
| `print_options` | PrintOptions | Print options |
| `protection` | SheetProtection | Sheet protection |
| `views` | SheetViewList | View settings |

### Cell Access

```python
# Direct coordinate access
ws['A1'] = 42
value = ws['A1'].value

# By row/column (1-based)
cell = ws.cell(row=1, column=1, value=42)

# Ranges
cells = ws['A1:C3']
column = ws['A']
row = ws[1]
```

### Iteration

```python
# Iterate rows
for row in ws.iter_rows(min_row=1, max_row=10, min_col=1, max_col=5):
    for cell in row:
        print(cell.value)

# Iterate columns
for col in ws.iter_cols(min_col=1, max_col=5):
    for cell in col:
        print(cell.value)

# Values only (no Cell objects)
for row in ws.iter_rows(values_only=True):
    print(row)

# All rows/columns
ws.rows      # Generator of row tuples
ws.columns   # Generator of column tuples
ws.values    # Generator of value tuples
```

### Modification Methods

| Method | Description |
|--------|-------------|
| `append(iterable)` | Append row of values |
| `insert_rows(idx, amount=1)` | Insert empty rows |
| `insert_cols(idx, amount=1)` | Insert empty columns |
| `delete_rows(idx, amount=1)` | Delete rows |
| `delete_cols(idx, amount=1)` | Delete columns |
| `move_range(range, rows=0, cols=0)` | Move cell range |
| `merge_cells(range_string)` | Merge cells |
| `unmerge_cells(range_string)` | Unmerge cells |

### Adding Objects

| Method | Description |
|--------|-------------|
| `add_chart(chart, anchor)` | Add chart at anchor cell |
| `add_image(image, anchor)` | Add image at anchor cell |
| `add_table(table)` | Add Excel table |
| `add_data_validation(dv)` | Add data validation |

---

## 4. Cell (`openpyxl.cell.cell.Cell`)

### Key Properties

| Property | Type | Description |
|----------|------|-------------|
| `value` | any | Cell value (auto-typed) |
| `data_type` | str | Type code: 's', 'n', 'f', 'b', 'd', 'e' |
| `row` | int | Row number (1-based) |
| `column` | int | Column number (1-based) |
| `coordinate` | str | Address (e.g., 'A1') |
| `column_letter` | str | Column letter |
| `hyperlink` | Hyperlink | Cell hyperlink |
| `comment` | Comment | Cell comment |
| `is_date` | bool | Value is date |

### Style Properties

| Property | Type | Description |
|----------|------|-------------|
| `font` | Font | Font formatting |
| `fill` | Fill | Background fill |
| `border` | Border | Cell borders |
| `alignment` | Alignment | Text alignment |
| `protection` | Protection | Cell protection |
| `number_format` | str | Number format code |
| `style` | str | Named style name |

### Data Types

| Code | Constant | Description |
|------|----------|-------------|
| `'s'` | TYPE_STRING | String |
| `'n'` | TYPE_NUMERIC | Number |
| `'f'` | TYPE_FORMULA | Formula |
| `'b'` | TYPE_BOOL | Boolean |
| `'d'` | TYPE_DATE | Date (ISO format) |
| `'e'` | TYPE_ERROR | Error (#VALUE!, etc.) |

---

## 5. Styles (`openpyxl.styles`)

### Font

```python
from openpyxl.styles import Font

font = Font(
    name='Calibri',
    size=11,
    bold=False,
    italic=False,
    underline='none',  # 'single', 'double', etc.
    strike=False,
    color='FF000000',  # ARGB hex
)
cell.font = font
```

### Fill

```python
from openpyxl.styles import PatternFill, GradientFill

# Solid fill
fill = PatternFill(
    fill_type='solid',
    start_color='FFFF00',
    end_color='FFFF00',
)

# Gradient fill
fill = GradientFill(
    type='linear',
    stop=['FF0000', '00FF00'],
)
cell.fill = fill
```

### Border

```python
from openpyxl.styles import Border, Side

border = Border(
    left=Side(style='thin', color='000000'),
    right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'),
    bottom=Side(style='thin', color='000000'),
)
cell.border = border
```

Border styles: `'thin'`, `'medium'`, `'thick'`, `'double'`, `'dotted'`, `'dashed'`, etc.

### Alignment

```python
from openpyxl.styles import Alignment

alignment = Alignment(
    horizontal='center',  # 'left', 'right', 'center', 'justify'
    vertical='center',    # 'top', 'center', 'bottom'
    wrap_text=True,
    text_rotation=0,      # 0-180 degrees
    shrink_to_fit=False,
    indent=0,
)
cell.alignment = alignment
```

### Protection

```python
from openpyxl.styles import Protection

protection = Protection(
    locked=True,
    hidden=False,
)
cell.protection = protection
```

### Named Styles

```python
from openpyxl.styles import NamedStyle, Font, Border

style = NamedStyle(name='highlight')
style.font = Font(bold=True, size=14)
style.border = Border(...)

wb.add_named_style(style)
cell.style = 'highlight'
```

### Number Formats

```python
cell.number_format = 'General'
cell.number_format = '0.00'
cell.number_format = '#,##0'
cell.number_format = '0%'
cell.number_format = 'YYYY-MM-DD'
cell.number_format = '$#,##0.00'
```

---

## 6. Charts (`openpyxl.chart`)

### Available Chart Types

| Class | Description |
|-------|-------------|
| `AreaChart`, `AreaChart3D` | Area charts |
| `BarChart`, `BarChart3D` | Bar/column charts |
| `BubbleChart` | Bubble charts |
| `DoughnutChart` | Doughnut charts |
| `LineChart`, `LineChart3D` | Line charts |
| `PieChart`, `PieChart3D` | Pie charts |
| `ProjectedPieChart` | Pie of pie / bar of pie |
| `RadarChart` | Radar charts |
| `ScatterChart` | XY scatter charts |
| `StockChart` | Stock (OHLC) charts |
| `SurfaceChart`, `SurfaceChart3D` | Surface charts |

### Creating Charts

```python
from openpyxl.chart import BarChart, Reference

# Create chart
chart = BarChart()
chart.title = "Sales Data"
chart.style = 10
chart.type = "col"  # or "bar" for horizontal

# Define data range
data = Reference(ws, min_col=2, min_row=1, max_col=4, max_row=10)
categories = Reference(ws, min_col=1, min_row=2, max_row=10)

chart.add_data(data, titles_from_data=True)
chart.set_categories(categories)

# Add to worksheet
ws.add_chart(chart, "F2")
```

### Chart Properties

| Property | Description |
|----------|-------------|
| `title` | Chart title |
| `style` | Built-in style number (1-48) |
| `type` | Chart subtype |
| `legend` | Legend settings |
| `x_axis`, `y_axis` | Axis settings |
| `width`, `height` | Dimensions (in cm) |

---

## 7. Images (`openpyxl.drawing.image`)

```python
from openpyxl.drawing.image import Image

img = Image('logo.png')
img.width = 100   # pixels
img.height = 100

ws.add_image(img, 'A1')
```

---

## 8. Comments (`openpyxl.comments`)

### Legacy Comments

```python
from openpyxl.comments import Comment

comment = Comment('This is a comment', 'Author Name')
comment.width = 300   # pixels
comment.height = 50

cell.comment = comment
```

### Reading Comments

```python
if cell.comment:
    print(cell.comment.text)
    print(cell.comment.author)
```

---

## 9. Data Validation (`openpyxl.worksheet.datavalidation`)

```python
from openpyxl.worksheet.datavalidation import DataValidation

# List validation
dv = DataValidation(
    type='list',
    formula1='"Option1,Option2,Option3"',
    allow_blank=True,
)
dv.error = 'Invalid entry'
dv.errorTitle = 'Error'
dv.prompt = 'Select from list'
dv.promptTitle = 'Options'

dv.add('A1:A100')
ws.add_data_validation(dv)

# Whole number validation
dv = DataValidation(type='whole', operator='between', formula1=1, formula2=100)

# Date validation
dv = DataValidation(type='date', operator='greaterThan', formula1='2020-01-01')
```

### Validation Types

`'whole'`, `'decimal'`, `'list'`, `'date'`, `'time'`, `'textLength'`, `'custom'`

### Operators

`'between'`, `'notBetween'`, `'equal'`, `'notEqual'`, `'lessThan'`, `'lessThanOrEqual'`, `'greaterThan'`, `'greaterThanOrEqual'`

---

## 10. Tables (`openpyxl.worksheet.table`)

```python
from openpyxl.worksheet.table import Table, TableStyleInfo

table = Table(displayName='SalesTable', ref='A1:D10')

style = TableStyleInfo(
    name='TableStyleMedium9',
    showFirstColumn=False,
    showLastColumn=False,
    showRowStripes=True,
    showColumnStripes=False,
)
table.tableStyleInfo = style

ws.add_table(table)
```

### Accessing Tables

```python
# List tables
for table_name in ws.tables:
    table = ws.tables[table_name]
    print(table.displayName, table.ref)
```

---

## 11. Conditional Formatting (`openpyxl.formatting`)

```python
from openpyxl.formatting.rule import (
    ColorScaleRule, FormulaRule, CellIsRule, Rule
)
from openpyxl.styles import PatternFill

# Color scale (2-color)
rule = ColorScaleRule(
    start_type='min', start_color='FF0000',
    end_type='max', end_color='00FF00',
)
ws.conditional_formatting.add('A1:A100', rule)

# Cell value rule
red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
rule = CellIsRule(operator='greaterThan', formula=['100'], fill=red_fill)
ws.conditional_formatting.add('B1:B100', rule)

# Formula rule
rule = FormulaRule(formula=['$A1>100'], fill=red_fill)
ws.conditional_formatting.add('A1:Z100', rule)
```

---

## 12. Page Setup and Printing

```python
# Page setup
ws.page_setup.orientation = 'landscape'  # or 'portrait'
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0

# Page margins (inches)
ws.page_margins.left = 0.5
ws.page_margins.right = 0.5
ws.page_margins.top = 0.75
ws.page_margins.bottom = 0.75

# Print options
ws.print_options.horizontalCentered = True
ws.print_options.verticalCentered = True

# Print area and titles
ws.print_area = 'A1:Z100'
ws.print_title_rows = '1:2'     # Repeat rows
ws.print_title_cols = 'A:B'     # Repeat columns

# Page breaks
ws.row_breaks.append(Break(id=20))  # Break after row 20
ws.col_breaks.append(Break(id=5))   # Break after column E
```

---

## 13. Utilities (`openpyxl.utils`)

### Cell Reference Utilities

```python
from openpyxl.utils import (
    get_column_letter,
    column_index_from_string,
    coordinate_to_tuple,
    range_boundaries,
)

get_column_letter(1)           # 'A'
get_column_letter(27)          # 'AA'
column_index_from_string('A')  # 1
column_index_from_string('AA') # 27

coordinate_to_tuple('A1')      # (1, 1)
range_boundaries('A1:C3')      # (1, 1, 3, 3)
```

### Formula Functions

```python
from openpyxl.utils import FORMULAE

'SUM' in FORMULAE    # True
'XLOOKUP' in FORMULAE  # True (Excel 365 functions included)
```

---

## 14. Read-Only Mode

For large files, use read-only mode to stream data without loading everything into memory:

```python
wb = load_workbook('large_file.xlsx', read_only=True)
ws = wb.active

for row in ws.iter_rows():
    for cell in row:
        print(cell.value)

wb.close()  # Important: close when done
```

**Limitations:**
- No random access to cells
- No modification
- Must iterate sequentially
- Must close workbook when done

---

## 15. Write-Only Mode

For creating large files efficiently:

```python
wb = Workbook(write_only=True)
ws = wb.create_sheet()

for row_data in data_generator():
    ws.append(row_data)

wb.save('large_file.xlsx')
```

**Limitations:**
- Can only append rows (no random cell access)
- Cannot read cells after writing
- No cell formatting via cell objects (use WriteOnlyCell)

```python
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font

cell = WriteOnlyCell(ws, value='Hello')
cell.font = Font(bold=True)
ws.append([cell, 'World'])
```

---

## 16. Document Properties

```python
from openpyxl import Workbook

wb = Workbook()

# Core properties
wb.properties.title = 'My Workbook'
wb.properties.subject = 'Data Analysis'
wb.properties.creator = 'John Doe'
wb.properties.description = 'Quarterly report'
wb.properties.keywords = 'sales, Q1, 2024'
wb.properties.category = 'Reports'

# Read properties
print(wb.properties.created)     # Creation date
print(wb.properties.modified)    # Last modified
print(wb.properties.lastModifiedBy)
```

---

## 17. Defined Names (Named Ranges)

```python
from openpyxl.workbook.defined_name import DefinedName

# Create named range
wb.defined_names['SalesData'] = DefinedName('SalesData', attr_text='Sheet1!$A$1:$D$100')

# Access named range
name = wb.defined_names['SalesData']
print(name.value)  # 'Sheet1!$A$1:$D$100'

# Iterate all names
for name in wb.defined_names:
    print(name.name, name.value)

# Delete
del wb.defined_names['SalesData']
```

---

## 18. Protection

### Workbook Protection

```python
wb.security.workbookPassword = 'secret'
wb.security.lockStructure = True
wb.security.lockWindows = True
```

### Sheet Protection

```python
ws.protection.sheet = True
ws.protection.password = 'secret'
ws.protection.enable()

# Specific permissions
ws.protection.formatCells = False
ws.protection.formatColumns = False
ws.protection.insertRows = True
ws.protection.deleteRows = True
ws.protection.sort = True
ws.protection.autoFilter = True
```

---

## 19. Merged Cells

```python
# Merge
ws.merge_cells('A1:D1')
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)

# Unmerge
ws.unmerge_cells('A1:D1')

# Check merged ranges
for merged_range in ws.merged_cells.ranges:
    print(merged_range)
```

**Note:** Only the top-left cell of a merged range holds the value.

---

## 20. Freeze Panes

```python
# Freeze rows above and columns left of cell
ws.freeze_panes = 'B2'   # Freeze row 1 and column A

ws.freeze_panes = 'A2'   # Freeze row 1 only
ws.freeze_panes = 'B1'   # Freeze column A only
ws.freeze_panes = None   # Unfreeze
```

---

## 21. Auto-Filter

```python
ws.auto_filter.ref = 'A1:D100'

# Add filter criteria
ws.auto_filter.add_filter_column(0, ['Value1', 'Value2'])
ws.auto_filter.add_sort_condition('B1:B100')
```

---

## 22. Hyperlinks

```python
from openpyxl.worksheet.hyperlink import Hyperlink

# Simple URL
ws['A1'].hyperlink = 'https://example.com'
ws['A1'].value = 'Click here'
ws['A1'].style = 'Hyperlink'  # Apply hyperlink style

# Internal link
ws['A2'].hyperlink = '#Sheet2!A1'

# Email
ws['A3'].hyperlink = 'mailto:user@example.com'
```

---

## 23. Rich Text (CellRichText)

```python
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

# Load with rich_text=True
wb = load_workbook('file.xlsx', rich_text=True)

# Create rich text
rich = CellRichText(
    'Normal text ',
    TextBlock(InlineFont(b=True), 'Bold text'),
    ' more normal',
)
ws['A1'] = rich
```

---

## Summary: Import Cheat Sheet

```python
# Core
from openpyxl import Workbook, load_workbook

# Styles
from openpyxl.styles import (
    Font, PatternFill, GradientFill, Border, Side,
    Alignment, Protection, NamedStyle, Color,
)

# Charts
from openpyxl.chart import (
    BarChart, LineChart, PieChart, AreaChart, ScatterChart,
    Reference, Series,
)

# Drawing
from openpyxl.drawing.image import Image

# Comments
from openpyxl.comments import Comment

# Data validation
from openpyxl.worksheet.datavalidation import DataValidation

# Tables
from openpyxl.worksheet.table import Table, TableStyleInfo

# Conditional formatting
from openpyxl.formatting.rule import (
    ColorScaleRule, FormulaRule, CellIsRule,
)

# Utilities
from openpyxl.utils import (
    get_column_letter, column_index_from_string,
    coordinate_to_tuple, range_boundaries,
)

# Write-only cells
from openpyxl.cell import WriteOnlyCell
```
