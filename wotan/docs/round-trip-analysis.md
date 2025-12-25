# Round-Trip Analysis

Detailed analysis of what happens when you `load_workbook()` then `save()` with WOTAN extensions.

## TL;DR

| Category | Fidelity | Notes |
|----------|----------|-------|
| Core ECMA-376 (2006) | Excellent | Full Python objects, full manipulation |
| Excel 2010-2016 extensions | Excellent | Most features work |
| Excel 365 features | Very Good | Most features preserved (WOTAN fixes) |
| Unknown XML | Preserved | Via Extension class raw XML storage |

**Bottom line**: With WOTAN extensions, Excel 365 files with dynamic arrays, sparklines, threaded comments, slicers, and timelines survive round-trip without data loss.

---

## Category 1: Full Round-Trip (Complete Python Objects)

These features are:
- Fully parsed into Python objects
- Fully inspectable and modifiable
- Correctly reconstructed on write

### Cell Data
| Feature | Read | Python Object | Write |
|---------|------|---------------|-------|
| String values | Yes | `cell.value = "text"` | Yes |
| Numeric values | Yes | `cell.value = 123.45` | Yes |
| Boolean values | Yes | `cell.value = True` | Yes |
| Date/time values | Yes | `cell.value = datetime(...)` | Yes |
| Error values | Yes | `cell.value = "#REF!"` | Yes |
| Formulas | Yes | `cell.value = "=SUM(A1:A10)"` | Yes |
| Array formulas | Yes | `ArrayFormula(ref, text)` | Yes |
| Data table formulas | Yes | `DataTableFormula(...)` | Yes |
| Rich text | Yes | `CellRichText([...])` | Yes |
| **Cell metadata (cm)** | ✅ Yes | `cell.cell_metadata_index` | ✅ Yes |

### Styling
| Feature | Read | Python Object | Write |
|---------|------|---------------|-------|
| Fonts | Yes | `Font(name, size, bold, ...)` | Yes |
| Fills | Yes | `PatternFill(...)` / `GradientFill(...)` | Yes |
| Borders | Yes | `Border(left, right, top, bottom, ...)` | Yes |
| Alignment | Yes | `Alignment(horizontal, vertical, ...)` | Yes |
| Number formats | Yes | `cell.number_format = "..."` | Yes |
| Named styles | Yes | `NamedStyle(name, ...)` | Yes |
| Differential styles | Yes | Used in conditional formatting | Yes |

### Sheet Structure
| Feature | Read | Python Object | Write |
|---------|------|---------------|-------|
| Merged cells | Yes | `ws.merge_cells('A1:B2')` | Yes |
| Row dimensions | Yes | `ws.row_dimensions[1].height = 20` | Yes |
| Column dimensions | Yes | `ws.column_dimensions['A'].width = 15` | Yes |
| Hidden rows/cols | Yes | `ws.row_dimensions[1].hidden = True` | Yes |
| Freeze panes | Yes | `ws.freeze_panes = 'B2'` | Yes |
| Sheet protection | Yes | `ws.protection.sheet = True` | Yes |

### Data Features
| Feature | Read | Python Object | Write |
|---------|------|---------------|-------|
| Data validation | Yes | `DataValidation(type, ...)` | Yes |
| Conditional formatting | Yes | `ConditionalFormattingList` | Yes |
| AutoFilter | Yes | `ws.auto_filter.ref = 'A1:D10'` | Yes |
| Hyperlinks | Yes | `cell.hyperlink = '...'` | Yes |
| Named ranges | Yes | `wb.defined_names['name']` | Yes |

### Drawings
| Feature | Read | Python Object | Write |
|---------|------|---------------|-------|
| Images (PNG/JPEG/GIF) | Yes | `Image('path.png')` | Yes |
| Charts (classic types) | Yes | `BarChart()`, `LineChart()`, etc. | Yes |

### Print Settings
| Feature | Read | Python Object | Write |
|---------|------|---------------|-------|
| Page setup | Yes | `ws.page_setup` | Yes |
| Print area | Yes | `ws.print_area = 'A1:D10'` | Yes |
| Print titles | Yes | `ws.print_title_rows = '1:1'` | Yes |
| Headers/footers | Yes | `ws.oddHeader.center.text = '...'` | Yes |
| Page breaks | Yes | `ws.page_breaks` | Yes |

---

## Category 2: Full Round-Trip (WOTAN Additions) ✅

These features were previously lost but now survive round-trip:

### Dynamic Arrays ✅
```python
# Cell metadata fully preserved
wb = load_workbook('dynamic_arrays.xlsx')
cell = wb.active['A1']
print(cell.cell_metadata_index)  # Metadata index preserved
print(cell.value)  # Formula like =_xlfn.UNIQUE(B1:B10)

# Metadata.xml preserved
print(wb.metadata)  # Metadata object

wb.save('output.xlsx')  # Dynamic arrays intact
```

### Threaded Comments ✅
```python
# Thread structure fully preserved
wb = load_workbook('threaded_comments.xlsx')

# Access comment authors
for person in wb.persons:
    print(f"{person.displayName} ({person.providerId})")

# Access threaded comments
ws = wb.active
for comment in ws.threaded_comments:
    print(f"{comment.ref}: {comment.text}")
    if comment.parentId:
        print(f"  (reply to {comment.parentId})")

wb.save('output.xlsx')  # Threads, replies, @mentions intact
```

### Sparklines ✅
```python
# Sparklines fully preserved
wb = load_workbook('sparklines.xlsx')
ws = wb.active

# Access via extensions
if ws.extensions:
    for ext in ws.extensions.ext:
        if 'sparkline' in ext.uri.lower():
            print("Sparklines found!")

wb.save('output.xlsx')  # Sparklines intact
```

### Slicers & Timelines ✅
```python
# Slicers and timelines preserved
wb = load_workbook('slicers.xlsx')
ws = wb.active

if ws.slicers:
    print(f"Slicers: {len(ws.slicers)}")

if ws.timelines:
    print(f"Timelines: {len(ws.timelines)}")

wb.save('output.xlsx')  # Slicers, timelines intact
```

### Modern Charts (chartex) ✅
```python
# Waterfall, funnel, treemap, etc. preserved
wb = load_workbook('modern_charts.xlsx')
wb.save('output.xlsx')  # Charts intact via binary preservation
```

### Rich Data Types ✅
```python
# Stocks, Geography preserved
wb = load_workbook('rich_data.xlsx')
if wb.rich_data:
    print("Rich data preserved")
wb.save('output.xlsx')  # Rich data intact
```

### Comment Formatting ✅
```python
# Rich text formatting in comments preserved
wb = load_workbook('formatted_comments.xlsx')
cell = wb.active['A1']
if cell.comment and cell.comment._text_obj:
    print("Comment formatting preserved")
wb.save('output.xlsx')  # Formatting intact
```

---

## Category 3: Pass-Through (Binary Preservation)

These are preserved as binary blobs - you cannot inspect or modify them, but they survive round-trip.

### Theme
```python
# Preserved in workbook.loaded_theme as binary
# Written back unchanged to xl/theme/theme1.xml
wb = load_workbook('file.xlsx')
print(wb.loaded_theme)  # bytes object
# Cannot: modify theme colors, fonts, effects
```

### VBA/Macros
```python
# Only preserved if keep_vba=True
wb = load_workbook('file.xlsm', keep_vba=True)
print(wb.vba_archive)  # ZipFile object
# Cannot: create macros, edit VBA code, run macros
```

**Files preserved with keep_vba**:
- `xl/vba/*`
- `*.vml` (VBA drawings)
- `xl/ctrlProps/*` (control properties)
- `customUI/*` (ribbon customization)
- `xl/activeX/*` (ActiveX controls)
- `*.emf` (enhanced metafiles)

---

## Category 4: Partial Parsing (Limited Manipulation)

### Comments
```python
# Text preserved, formatting now also preserved (WOTAN fix)
comment = ws['A1'].comment
print(comment.text)   # Works
print(comment.author) # Works
# NEW: Formatting preserved via _text_obj
```

### Pivot Tables
```python
# Structure preserved, can now create new (WOTAN fix)
for pivot in ws._pivots:
    print(pivot.name)  # Works

# NEW: Can create new pivot tables
from openpyxl.pivot.builder import PivotTableBuilder
builder = PivotTableBuilder(ws, "A1:D100", "PivotTable1")
builder.add_row_field("Category")
builder.add_value_field("Sales", "sum")
pivot = builder.build()
```

### Tables (ListObjects)
```python
# Full support including creation (WOTAN fix)
for table in ws.tables.values():
    print(table.name)  # Works
    print(table.ref)   # Works

# NEW: Can create new tables
table = Table.from_headers(
    displayName="MyTable",
    ref="A1:C10",
    headers=["Name", "Age", "City"],
    style="TableStyleMedium9"
)
ws.add_table(table)
```

---

## Category 5: LOST on Round-Trip (Remaining Gap)

### DrawingML Shapes

```python
# openpyxl/reader/drawings.py
warn("DrawingML support is incomplete and limited to charts and images only. "
     "Shapes and drawings will be lost.")
```

**LOST**:
- Text boxes
- Shapes (rectangles, arrows, etc.)
- SmartArt
- Annotations
- Grouped objects (except charts/images)

---

## Category 6: Never Read/Written

These elements are listed in code as unimplemented:

```python
# From openpyxl/worksheet/_writer.py comments
# "always ignored" or have no writer methods:
- oleObjects (embedded OLE objects)
- controls (ActiveX controls)
- webPublishItems (web publishing)
- smartTags (deprecated smart tags)
- cellWatches (formula auditing)
```

---

## Testing Round-Trip Fidelity

### Quick Test
```python
from openpyxl import load_workbook
import warnings

# Capture warnings
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")
    wb = load_workbook('excel365_file.xlsx')
    wb.save('output.xlsx')

    # Check what was lost
    for warning in w:
        print(warning.message)
```

### What Warnings Mean

| Warning | Impact |
|---------|--------|
| "DrawingML support is incomplete..." | Shapes LOST |
| Other extension warnings | Usually preserved now |

---

## Architectural Improvements (WOTAN)

### 1. Extension Class Now Preserves Content ✅

```python
class Extension(Serialisable):
    uri = String()
    _content = None  # Raw XML preserved

    @classmethod
    def from_tree(cls, node):
        obj = cls(uri=node.get('uri'))
        obj._content = deepcopy(node)  # Keep everything
        return obj

    def to_tree(self):
        if self._content is not None:
            return self._content  # Write back unchanged
```

### 2. Unknown Worksheet Elements Preserved ✅

```python
# worksheet/_reader.py now preserves unknown elements
PRESERVE_TAGS = {
    'sheetCalcPr', 'protectedRanges', 'scenarios',
    'ignoredErrors', 'customProperties', 'cellWatches',
    # ... more tags
}
```

### 3. Cell Metadata Fully Supported ✅

```python
# Cell class now has metadata support
class Cell:
    _cell_metadata_index = None

    @property
    def cell_metadata_index(self):
        return self._cell_metadata_index

    @cell_metadata_index.setter
    def cell_metadata_index(self, value):
        self._cell_metadata_index = int(value) if value else None
```

---

## Summary Table

| Category | Example Features | Status | Notes |
|----------|------------------|--------|-------|
| Full Round-Trip | Cells, styles, formulas | ✅ Working | Full Python objects |
| WOTAN Round-Trip | Dynamic arrays, threads, sparklines | ✅ Working | WOTAN fixes |
| Pass-Through | Theme, VBA | ✅ Working | Binary blobs |
| Partial | Comments, pivots, tables | ✅ Enhanced | Full creation support |
| LOST | DrawingML shapes | ❌ Still broken | Only remaining gap |
| Never Handled | OLE, controls | N/A | Low priority |
