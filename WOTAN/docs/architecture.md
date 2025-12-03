# Openpyxl Architecture

This document describes the existing openpyxl architecture and identifies extension points for WOTAN work.

## Overview

Openpyxl is organized into ~23 main modules with ~190 Python files. It uses a descriptor-based system for XML serialization and follows the OOXML package structure.

## Directory Structure

```
openpyxl/
├── cell/                  # Cell types and rich text
│   ├── cell.py           # Main Cell class
│   ├── rich_text.py      # CellRichText for multi-format text
│   ├── text.py           # Text elements
│   └── _writer.py        # Cell XML writer
│
├── chart/                 # Chart support (37 modules)
│   ├── _chart.py         # Base chart classes
│   ├── area_chart.py     # Area charts
│   ├── bar_chart.py      # Bar/column charts
│   ├── line_chart.py     # Line charts
│   ├── pie_chart.py      # Pie/doughnut charts
│   ├── scatter_chart.py  # XY scatter plots
│   └── ...               # Many more chart types
│
├── chartsheet/            # Chart-only sheets
│
├── comments/              # Cell comments
│   ├── comments.py       # Comment class
│   └── comment_sheet.py  # Comment collections
│
├── compat/                # Python compatibility
│
├── descriptors/           # XML serialization descriptors
│   ├── base.py           # Base descriptor classes
│   ├── serialisable.py   # Serialisable base class
│   ├── sequence.py       # Sequence descriptors
│   └── nested.py         # Nested element handling
│
├── drawing/               # Drawings, images, shapes
│   ├── image.py          # Image handling
│   ├── spreadsheet_drawing.py  # Drawing container
│   └── ...               # Shapes, effects, etc.
│
├── formatting/            # Conditional formatting
│   ├── rule.py           # Formatting rules
│   └── formatting.py     # Rule collections
│
├── formula/               # Formula handling
│   ├── tokenizer.py      # Excel formula tokenizer
│   └── translate.py      # Formula translation
│
├── packaging/             # OOXML package management
│   ├── core.py           # Core properties
│   ├── custom.py         # Custom properties
│   ├── manifest.py       # Content types
│   ├── relationship.py   # Relationships
│   └── workbook.py       # Workbook package info
│
├── pivot/                 # Pivot tables
│   ├── table.py          # Pivot table definition
│   ├── cache.py          # Pivot cache
│   └── fields.py         # Pivot fields
│
├── reader/                # File reading
│   ├── excel.py          # Main workbook reader
│   └── workbook.py       # Workbook parsing
│
├── styles/                # Cell styling
│   ├── fonts.py          # Font styling
│   ├── fills.py          # Fill patterns
│   ├── borders.py        # Border styles
│   ├── alignment.py      # Text alignment
│   ├── colors.py         # Color handling
│   └── named_styles.py   # Named style management
│
├── utils/                 # Utility functions
│   ├── formulas.py       # Formula validation & FORMULAE list
│   ├── cell.py           # Cell utilities
│   └── ...
│
├── workbook/              # Workbook management
│   ├── workbook.py       # Main Workbook class
│   ├── defined_name.py   # Named ranges
│   ├── external_link/    # External references
│   ├── protection.py     # Workbook protection
│   ├── properties.py     # Calculation properties
│   ├── views.py          # Workbook views
│   └── _writer.py        # Workbook writer
│
├── worksheet/             # Worksheet implementation (35 modules)
│   ├── worksheet.py      # Main Worksheet class
│   ├── _reader.py        # Worksheet reader
│   ├── _writer.py        # Worksheet writer
│   ├── cell_range.py     # Cell range handling
│   ├── datavalidation.py # Data validation
│   ├── filters.py        # AutoFilter
│   ├── formula.py        # ArrayFormula, DataTableFormula
│   ├── hyperlink.py      # Hyperlinks
│   ├── merge.py          # Merged cells
│   ├── page.py           # Page layout
│   ├── pagebreak.py      # Page breaks
│   ├── print_settings.py # Print settings
│   ├── properties.py     # Sheet properties
│   ├── protection.py     # Sheet protection
│   ├── table.py          # Table definitions
│   ├── views.py          # Sheet views
│   └── ...
│
├── writer/                # File writing
│   └── excel.py          # Main workbook writer
│
└── xml/                   # XML utilities
    ├── constants.py      # Namespaces, paths, content types
    └── functions.py      # XML helper functions
```

## Core Patterns

### 1. Descriptor-Based Serialization

Openpyxl uses Python descriptors for XML attribute mapping. This is the foundation for all XML handling.

```python
# openpyxl/descriptors/base.py
class Descriptor:
    """Base class for all descriptors"""

class String(Descriptor):
    """String attribute"""

class Integer(Descriptor):
    """Integer attribute"""

class Bool(Descriptor):
    """Boolean attribute"""

class Typed(Descriptor):
    """Typed object attribute"""
```

**Usage pattern**:
```python
from openpyxl.descriptors import Typed, String, Bool
from openpyxl.descriptors.serialisable import Serialisable

class MyElement(Serialisable):
    tagname = "myElement"

    name = String()
    enabled = Bool(allow_none=True)
    child = Typed(expected_type=ChildElement, allow_none=True)

    def __init__(self, name=None, enabled=None, child=None):
        self.name = name
        self.enabled = enabled
        self.child = child
```

### 2. Serialisable Base Class

All XML-serializable classes inherit from `Serialisable`:

```python
# openpyxl/descriptors/serialisable.py
class Serialisable:
    tagname = None  # XML tag name
    namespace = None  # Optional namespace

    @classmethod
    def from_tree(cls, node):
        """Deserialize from XML element"""

    def to_tree(self, tagname=None, idx=None, namespace=None):
        """Serialize to XML element"""
```

### 3. Namespace Management

Namespaces are centralized in `openpyxl/xml/constants.py`:

```python
# Core namespaces
SHEET_MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Content types
WORKSHEET_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

# Extension GUIDs
EXT_TYPES = {
    '{05C60535-1F16-4FD2-B633-F4F36F0B64E0}': 'Sparkline Group',
    # ...
}
```

### 4. Relationship Management

OOXML uses relationships to link parts:

```python
# openpyxl/packaging/relationship.py
class Relationship:
    Type = String()
    Target = String()
    TargetMode = String(allow_none=True)
    Id = String(allow_none=True)
```

### 5. Extension Lists (extLst)

Many elements support extension lists for future features:

```python
from openpyxl.descriptors.excel import ExtensionList

class SomeElement(Serialisable):
    # ... regular attributes ...
    extLst = Typed(expected_type=ExtensionList, allow_none=True)
```

---

## Key Extension Points for WOTAN

### 1. Adding New Namespaces

**File**: `openpyxl/xml/constants.py`

```python
# Add after existing namespaces (~line 75)
X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
X15_NS = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
XLDAPR_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"
XLRD_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
XLTHREADED_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
```

### 2. Adding New Content Types

**File**: `openpyxl/xml/constants.py`

```python
# Add after existing MIME types (~line 110)
METADATA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
THREADED_COMMENTS_TYPE = "application/vnd.ms-excel.threadedcomments+xml"
PERSON_TYPE = "application/vnd.ms-excel.person+xml"
RICHVALUE_TYPE = "application/vnd.ms-excel.rdRichValue+xml"
```

### 3. Adding Cell Metadata

**File**: `openpyxl/cell/_writer.py`

The cell writer needs to add `cm` attribute for dynamic arrays:

```python
# Current cell writing (simplified)
def write_cell(xf, worksheet, cell, styled=None):
    # ... existing code ...
    el.set('r', cell.coordinate)
    if cell.data_type != 'f':
        el.set('t', cell.data_type)
    # WOTAN: Add cell metadata index
    # if cell.metadata_index is not None:
    #     el.set('cm', str(cell.metadata_index))
```

### 4. Adding Workbook Parts

**File**: `openpyxl/workbook/_writer.py`

Add relationships for new parts:

```python
def write_workbook_rels(workbook):
    # ... existing relationships ...
    # WOTAN: Add metadata relationship
    # if workbook.has_metadata:
    #     rel = Relationship(Type="metadata", Target="metadata.xml")
```

### 5. Creating New Part Handlers

Pattern for new XML parts (e.g., `metadata.xml`):

```python
# openpyxl/packaging/metadata.py (NEW)
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import Typed, Sequence
from openpyxl.xml.constants import XLDAPR_NS

class FutureMetadata(Serialisable):
    tagname = "futureMetadata"
    namespace = XLDAPR_NS

    # ... attributes ...

class MetadataTypes(Serialisable):
    tagname = "metadataTypes"

    metadataType = Sequence(expected_type=MetadataType)

def read_metadata(archive, path):
    """Read metadata.xml from archive"""

def write_metadata(workbook):
    """Generate metadata.xml content"""
```

### 6. Extending Formula Support

**File**: `openpyxl/utils/formulas.py`

Add new functions to FORMULAE tuple:

```python
# Current (line 7)
FORMULAE = ("CUBEKPIMEMBER", "CUBEMEMBER", ...)

# WOTAN: Add modern functions
MODERN_FORMULAE = (
    "LAMBDA", "LET", "MAKEARRAY", "MAP", "REDUCE", "SCAN",
    "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY",
    "XLOOKUP", "XMATCH",
    "TEXTJOIN", "CONCAT", "TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT",
    "GROUPBY", "PIVOTBY", "PERCENTOF",
    "SWITCH", "IFS", "XOR", "IFNA", "MAXIFS", "MINIFS",
    "IMAGE", "FIELDVALUE",
)

FORMULAE = frozenset(FORMULAE + MODERN_FORMULAE)
```

**File**: `openpyxl/formula/tokenizer.py`

Handle spilled range operator:

```python
# Add to TOKEN_ENDERS or create handler
# The '#' operator marks spilled ranges: A1# means "spill from A1"
```

### 7. Adding Threaded Comments

Create new module structure:

```
openpyxl/comments/
├── __init__.py
├── comments.py        # Existing
├── comment_sheet.py   # Existing
├── threaded.py        # NEW - ThreadedComment class
└── person.py          # NEW - Person class
```

### 8. Adding Chart Types

**Location**: `openpyxl/chart/`

Follow existing patterns:

```python
# openpyxl/chart/waterfall_chart.py (NEW)
from openpyxl.chart._chart import ChartBase

class WaterfallChart(ChartBase):
    tagname = "waterfallChart"
    # ... implement per spec ...
```

---

## Reader/Writer Flow

### Reading an XLSX File

```
load_workbook(filename)
    ↓
openpyxl/reader/excel.py: ExcelReader
    ↓
├── Read [Content_Types].xml (manifest)
├── Read _rels/.rels (package relationships)
├── Read xl/_rels/workbook.xml.rels (workbook relationships)
├── Read xl/workbook.xml (workbook definition)
├── Read xl/styles.xml (styles)
├── Read xl/sharedStrings.xml (shared strings)
├── For each worksheet:
│   ├── Read xl/worksheets/sheet*.xml
│   └── Read xl/worksheets/_rels/sheet*.xml.rels
├── Read charts, drawings, etc.
└── Return Workbook object
```

### Writing an XLSX File

```
workbook.save(filename)
    ↓
openpyxl/writer/excel.py: ExcelWriter
    ↓
├── Create ZIP archive
├── Write [Content_Types].xml
├── Write _rels/.rels
├── Write docProps/core.xml
├── Write docProps/app.xml
├── Write xl/workbook.xml
├── Write xl/_rels/workbook.xml.rels
├── Write xl/styles.xml
├── Write xl/sharedStrings.xml
├── For each worksheet:
│   ├── Write xl/worksheets/sheet*.xml
│   └── Write xl/worksheets/_rels/sheet*.xml.rels
├── Write charts, drawings, etc.
└── Close archive
```

---

## Testing Patterns

### Unit Tests

```python
# openpyxl/*/tests/test_*.py
import pytest
from openpyxl.xxx import SomeClass

class TestSomeClass:
    def test_basic(self):
        obj = SomeClass(name="test")
        assert obj.name == "test"

    def test_from_tree(self):
        src = """<element name="test"/>"""
        node = fromstring(src)
        obj = SomeClass.from_tree(node)
        assert obj.name == "test"

    def test_to_tree(self):
        obj = SomeClass(name="test")
        tree = obj.to_tree()
        assert tree.get("name") == "test"
```

### Round-Trip Tests

```python
def test_roundtrip(datadir, tmp_path):
    datadir.join("test_file.xlsx").copy(tmp_path)
    wb = load_workbook(tmp_path / "test_file.xlsx")
    # ... modify ...
    wb.save(tmp_path / "output.xlsx")

    wb2 = load_workbook(tmp_path / "output.xlsx")
    # ... verify ...
```

### Test Data

- Test files go in `openpyxl/tests/data/`
- Create minimal XLSX files in Excel for specific features
- Use `@pytest.fixture` for common setup

---

## Key Classes Reference

| Class | Location | Purpose |
|-------|----------|---------|
| `Workbook` | `workbook/workbook.py` | Main workbook container |
| `Worksheet` | `worksheet/worksheet.py` | Sheet data and properties |
| `Cell` | `cell/cell.py` | Individual cell |
| `Serialisable` | `descriptors/serialisable.py` | XML serialization base |
| `Relationship` | `packaging/relationship.py` | OOXML relationships |
| `ArrayFormula` | `worksheet/formula.py` | Array formula representation |
| `Comment` | `comments/comments.py` | Cell comment |
| `Font`, `Fill`, `Border` | `styles/*.py` | Cell styling |
| `ChartBase` | `chart/_chart.py` | Chart base class |

---

## Adding a New Feature Checklist

1. [ ] Research: Read MS-XLSX spec for the feature
2. [ ] Test data: Create XLSX in Excel with the feature
3. [ ] Namespace: Add to `xml/constants.py` if needed
4. [ ] Content type: Add to `xml/constants.py` if needed
5. [ ] Classes: Create `Serialisable` subclasses for XML elements
6. [ ] Reader: Add parsing in appropriate reader module
7. [ ] Writer: Add serialization in appropriate writer module
8. [ ] Relationships: Update relationship handlers if needed
9. [ ] Unit tests: Test class serialization/deserialization
10. [ ] Round-trip tests: Test full read/modify/write cycle
11. [ ] Excel validation: Verify output opens correctly in Excel
