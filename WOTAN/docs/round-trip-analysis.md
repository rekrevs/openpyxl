# Round-Trip Analysis

Detailed analysis of what happens when you `load_workbook()` then `save()`.

## TL;DR

| Category | Fidelity | Notes |
|----------|----------|-------|
| Core ECMA-376 (2006) | Excellent | Full Python objects, full manipulation |
| Excel 2010-2016 extensions | Good | Most features work |
| Excel 365 features | Poor | Many features LOST |
| Unknown XML | None | Silently dropped |

**Bottom line**: If someone sends you an Excel 365 file with dynamic arrays, sparklines, or threaded comments, opening and saving it with openpyxl will **destroy those features**.

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

## Category 2: Pass-Through (Binary Preservation)

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

## Category 3: Partial Parsing (Limited Manipulation)

### Comments
```python
# Text is preserved, formatting is LOST
comment = ws['A1'].comment
print(comment.text)   # Works
print(comment.author) # Works
# LOST: font formatting, colors, dimensions
```

**What happens**:
- Read: Text and author extracted, formatting discarded
- Write: Plain text comment written
- Impact: Rich formatting in comments lost

### Pivot Tables
```python
# Structure preserved, cannot create new
for pivot in ws._pivots:
    print(pivot.name)  # Works
# Cannot: create new pivot table from scratch
```

**What happens**:
- Read: Pivot definition parsed
- Write: Existing definition written back
- Impact: Can modify some properties, cannot create new

### Tables (ListObjects)
```python
# Definition preserved, cannot create new
for table in ws.tables.values():
    print(table.name)  # Works
    print(table.ref)   # Works
# Cannot: create new table from scratch
```

---

## Category 4: LOST on Round-Trip (Critical)

### Extension List Content (extLst)

**The biggest problem.** Excel stores modern features in `<extLst>` elements.

```xml
<!-- What Excel writes -->
<extLst>
  <ext uri="{05C60535-1F16-4FD2-B633-F4F36F0B64E0}">
    <x14:sparklineGroups>
      <!-- Complex sparkline definition -->
    </x14:sparklineGroups>
  </ext>
</extLst>
```

```python
# What openpyxl does
class Extension(Serialisable):
    uri = String()  # Only this is kept!
    # Content inside <ext> is DISCARDED
```

**Result**:
- Warning: "Sparkline Group extension is not supported and will be removed"
- All sparklines gone from saved file

**Affected features**:
| Extension | GUID | Status |
|-----------|------|--------|
| Sparklines | `{05C60535-...}` | LOST |
| Slicers | `{A8765BA9-...}` | LOST |
| Timelines | `{7E03D99C-...}` | LOST |
| Protected ranges (ext) | `{FC87AEE6-...}` | LOST |
| Ignored errors | `{01252117-...}` | LOST |
| Web extensions | `{F7C9EE02-...}` | LOST |

### Dynamic Arrays

```xml
<!-- Excel 365 cell with dynamic array -->
<c r="A1" cm="1">  <!-- cm="1" marks dynamic array -->
  <f>_xlfn.UNIQUE(B1:B10)</f>
</c>
```

```python
# openpyxl ignores cm attribute
cell = ws['A1']
print(cell.value)  # Formula text preserved
# BUT: cm attribute not read, metadata.xml not parsed
# Result: Formula becomes static on save
```

**What's missing**:
- `cm` attribute on cells (cell metadata index)
- `metadata.xml` file handling
- `futureMetadata` with XLDAPR namespace
- Spill range operator (`#`) in formulas

### Threaded Comments

```xml
<!-- Excel 365 uses separate files -->
xl/threadedComments/threadedComment1.xml
xl/persons/person.xml
```

```python
# openpyxl doesn't read these files at all
# Only legacy comments (in xl/comments*.xml) are read
```

**Result**: Comment threads, replies, @mentions - all LOST

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

### Rich Data Types

```xml
<!-- Stocks, Geography, etc. -->
xl/richData/rdrichvalue.xml
xl/richData/rdRichValueTypes.xml
```

**Not parsed at all** - linked data types completely lost.

### Modern Chart Types

Some newer chart types throw errors and are dropped:
- Waterfall
- Funnel
- Treemap
- Sunburst
- Box & Whisker
- Histogram
- Map charts

### Unknown XML Elements

```python
# openpyxl/worksheet/_reader.py uses iterparse
for _, element in it:
    if tag_name in dispatcher:
        dispatcher[tag_name](element)
    # ELSE: Element silently dropped, no fallback
```

**Any element not in the dispatcher is gone forever.**

---

## Category 5: Never Read/Written

These elements are listed in code as unimplemented:

```python
# From openpyxl/worksheet/_writer.py comments
# "always ignored" or have no writer methods:
- oleObjects (embedded OLE objects)
- controls (ActiveX controls)
- webPublishItems (web publishing)
- smartTags (deprecated smart tags)
- cellWatches (formula auditing)
- ignoredErrors (error ignore flags)
- customProperties (sheet custom props)
- drawingHF (header/footer drawings)
- background (sheet background image)
- phonetic (phonetic text properties)
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
| "... extension is not supported and will be removed" | Feature LOST |
| "DrawingML support is incomplete..." | Shapes LOST |
| "Data Validation extension..." | Some validation rules LOST |
| "Unknown type ... in cell" | Cell type may not round-trip |

---

## Architectural Root Causes

### 1. Serialisable Pattern Drops Unknown

```python
class Serialisable:
    @classmethod
    def from_tree(cls, node):
        # Only processes defined descriptors
        # Unknown children silently ignored
```

### 2. No Raw XML Storage

There's no mechanism to store raw XML for later reconstruction:
```python
# Would need something like:
class Serialisable:
    _raw_xml = None  # Store original for unknown elements
```

### 3. Extension Class Too Simple

```python
# Current - only stores URI
class Extension(Serialisable):
    uri = String()

# Needed - stores full content
class Extension(Serialisable):
    uri = String()
    _content = None  # Raw XML element
```

### 4. Reader Uses Streaming Without Fallback

```python
# Worksheet reader clears elements as it goes
element.clear()  # Memory efficient, but data gone
```

---

## Recommendations for WOTAN

### Phase 0: Stop Losing Data

1. **Modify Extension class** to preserve raw XML content
2. **Add unknown element storage** to Serialisable base
3. **Create preservation mode** for readers
4. **Add round-trip tests** with Excel 365 files

### Phase 1: Measure the Problem

1. Create test files in Excel 365 with each modern feature
2. Run round-trip and document warnings
3. Compare input/output XML to catalog losses
4. Prioritize based on frequency of use

### Phase 2: Incremental Parsing

1. First preserve, then parse
2. Each feature gets full Python object model
3. Unknown features still preserved as raw XML
4. Never lose data we don't understand

---

## Summary Table

| Category | Example Features | Status | Action Needed |
|----------|------------------|--------|---------------|
| Full Round-Trip | Cells, styles, formulas | Working | None |
| Pass-Through | Theme, VBA | Working | None |
| Partial | Comments, pivots | Limited | Enhance |
| LOST | extLst, dynamic arrays | Broken | Critical fix |
| Never Handled | OLE, controls | Missing | Evaluate need |
