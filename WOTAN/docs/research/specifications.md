# WOTAN Implementation Specifications

Comprehensive specifications for all NEEDS-SPEC backlog items.

## Table of Contents

1. [P0: Preservation](#p0-preservation)
   - [B-PRESERVE-03: Serialisable Raw XML Storage](#b-preserve-03)
   - [B-PRESERVE-04: DrawingML Shape Preservation](#b-preserve-04)
2. [P1: Dynamic Arrays](#p1-dynamic-arrays)
   - [B-DYNARR-02: Cell Metadata Attribute](#b-dynarr-02)
   - [B-DYNARR-03: futureMetadata XLDAPR](#b-dynarr-03)
   - [B-DYNARR-04: Spilled Range Operator](#b-dynarr-04)
3. [P2: Comments & Charts](#p2-comments-charts)
   - [B-COMMENT-01: Threaded Comments Reader](#b-comment-01)
   - [B-COMMENT-02: Person.xml Support](#b-comment-02)
   - [B-CHART-01-06: Modern Chart Types](#b-chart-01-06)
4. [P3: Slicers, Rich Data, Sparklines](#p3-slicers-rich-sparklines)
   - [B-SLICER-01: Table Slicer Reader](#b-slicer-01)
   - [B-SLICER-02: Timeline Reader](#b-slicer-02)
   - [B-RICH-01: xlRichValue Reader](#b-rich-01)
   - [B-SPARK-01: Sparkline Reader](#b-spark-01)
5. [P4: Pivot Tables & Tables](#p4-pivot-tables)
   - [B-PIVOT-01: Pivot Table Creation](#b-pivot-01)
   - [B-TABLE-01: Table Creation](#b-table-01)

---

## P0: Preservation

### B-PRESERVE-03: Serialisable Raw XML Storage {#b-preserve-03}

**Intent**: Allow any Serialisable to preserve unknown child elements.

**Current State**: Unknown XML children are silently discarded in `from_tree()`.

**Solution**: Add opt-in `_preserve_unknown` class attribute.

**Files to Modify**:
- `openpyxl/descriptors/serialisable.py`

**Implementation**:

```python
class Serialisable(metaclass=MetaSerialisable):
    _preserve_unknown = False  # Class-level opt-in

    @classmethod
    def from_tree(cls, node):
        # ... existing code ...
        obj = cls(**attrib)

        # NEW: Preserve unknown children if enabled
        if getattr(cls, '_preserve_unknown', False):
            obj._unknown_children = []
            for el in node:
                tag = localname(el)
                if tag not in handled_tags:
                    obj._unknown_children.append(deepcopy(el))
        return obj

    def to_tree(self, ...):
        # ... existing code ...

        # NEW: Append preserved unknown children
        if getattr(self.__class__, '_preserve_unknown', False):
            for child in getattr(self, '_unknown_children', []):
                el.append(deepcopy(child))
        return el
```

**Effort**: Low (50 lines)

---

### B-PRESERVE-04: DrawingML Shape Preservation {#b-preserve-04}

**Intent**: Stop losing textboxes, shapes, annotations.

**Current State**: `find_images()` in reader crashes with TypeError, discards all shapes.

**Solution**: Create `RawShapeElement` class using Extension pattern.

**Files to Modify**:
- `openpyxl/drawing/connector.py` - Add RawShapeElement
- `openpyxl/drawing/spreadsheet_drawing.py` - Override anchor from_tree/to_tree

**Implementation**:

```python
# In connector.py
class RawShapeElement(Serialisable):
    """Preserves unparsed shape XML for round-trip fidelity"""
    tagname = "sp"

    def __init__(self):
        self._content = None

    @classmethod
    def from_tree(cls, node):
        obj = cls()
        obj._content = _convert_to_lxml(node)
        return obj

    def to_tree(self, ...):
        if self._content is not None:
            return deepcopy(self._content)
```

**Effort**: Medium (100 lines)

---

## P1: Dynamic Arrays

### B-DYNARR-02: Cell Metadata Attribute {#b-dynarr-02}

**Intent**: Add `cm` attribute to cells for dynamic array marking.

**Files to Modify**:
- `openpyxl/cell/cell.py` - Add `_cell_metadata_index` to `__slots__`
- `openpyxl/cell/_writer.py` - Write `cm` attribute
- `openpyxl/worksheet/_reader.py` - Read `cm` attribute

**Implementation**:

```python
# cell.py - Add to __slots__
__slots__ = (..., '_cell_metadata_index')

# cell.py - Add property
@property
def cell_metadata_index(self):
    return self._cell_metadata_index

# _writer.py - In _set_attributes()
if cell._cell_metadata_index is not None:
    attrs['cm'] = str(cell._cell_metadata_index)

# _reader.py - In parse_cell()
cell_metadata_index = element.get('cm')
if cell_metadata_index:
    cell_metadata_index = int(cell_metadata_index)
```

**Effort**: Low (30 lines)

---

### B-DYNARR-03: futureMetadata XLDAPR {#b-dynarr-03}

**Intent**: Implement dynamic array property metadata structure.

**Status**: Already implemented in B-DYNARR-01! The `FutureMetadata` and `FutureMetadataBlock` classes exist in `openpyxl/packaging/metadata.py`.

**Remaining Work**: None - preservation already works.

---

### B-DYNARR-04: Spilled Range Operator {#b-dynarr-04}

**Intent**: Parse and preserve the `#` operator in formulas (A1# syntax).

**Current State**: Tokenizer fails with "Unexpected character" for A1#.

**Files to Modify**:
- `openpyxl/formula/tokenizer.py`

**Implementation**:

```python
# Add '#' to TOKEN_ENDERS (line 43)
TOKEN_ENDERS = ',;}) +-*/^&=><%#'

# Modify _parse_error() to handle spilled ranges
def _parse_error(self):
    # Check if this follows a RANGE operand (spilled range)
    if (self.items and
        self.items[-1].type == Token.OPERAND and
        self.items[-1].subtype == Token.RANGE):
        self.items[-1].value += '#'
        return 1

    # Otherwise parse as error code (existing logic)
    ...
```

**Effort**: Low (20 lines)

---

## P2: Comments & Charts

### B-COMMENT-01: Threaded Comments Reader {#b-comment-01}

**Intent**: Read threaded comments from modern Excel files.

**Files to Create**:
- `openpyxl/comments/threaded.py`

**Classes**:

```python
class ThreadedComment(Serialisable):
    tagname = "threadedComment"

    ref = String()           # Cell reference
    dT = DateTime()          # Timestamp
    personId = String()      # GUID to person
    id = String()            # Comment GUID
    parentId = String()      # Parent for replies
    done = Bool()            # Resolved status
    text = Typed(expected_type=Text)

class ThreadedComments(Serialisable):
    tagname = "ThreadedComments"
    threadedComment = Sequence(expected_type=ThreadedComment)
```

**Namespace**: `http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments`

**Effort**: Medium (150 lines)

---

### B-COMMENT-02: Person.xml Support {#b-comment-02}

**Intent**: Read and write person data for threaded comments.

**Files to Create**:
- `openpyxl/comments/person.py`

**Classes**:

```python
class Person(Serialisable):
    tagname = "person"

    displayName = String()
    id = String()            # GUID (uppercase)
    userId = String()
    providerId = String()

class PersonList(Serialisable):
    tagname = "personList"
    person = Sequence(expected_type=Person)
```

**Content Type**: `application/vnd.ms-excel.person+xml`
**Relationship**: `http://schemas.microsoft.com/office/2017/10/relationships/person`

**Effort**: Medium (100 lines)

---

### B-CHART-01-06: Modern Chart Types {#b-chart-01-06}

**Intent**: Add waterfall, funnel, treemap, sunburst, box-whisker, histogram charts.

**Critical Finding**: These use a DIFFERENT architecture than existing charts!

| Standard Charts | Extended Charts (New) |
|----------------|----------------------|
| `application/vnd.openxmlformats-officedocument.drawingml.chart+xml` | `application/vnd.ms-office.chartex+xml` |
| `xmlns:c` namespace | `xmlns:cx` namespace |
| `<c:chartSpace>` | `<cx:chartSpace>` |

**Files to Create**:
- `openpyxl/chart/chartex.py` - New ChartExBase class
- `openpyxl/chart/waterfall_chart.py`
- `openpyxl/chart/funnel_chart.py`
- `openpyxl/chart/treemap_chart.py`
- `openpyxl/chart/sunburst_chart.py`
- `openpyxl/chart/boxwhisker_chart.py`
- `openpyxl/chart/histogram_chart.py`

**XML Structure**:
```xml
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chart>
    <cx:waterfallChart>
      <cx:seriesLayout val="waterfall"/>
      ...
    </cx:waterfallChart>
  </cx:chart>
</cx:chartSpace>
```

**Effort**: High (500+ lines, requires new base class)

---

## P3: Slicers, Rich Data, Sparklines

### B-SLICER-01: Table Slicer Reader {#b-slicer-01}

**Intent**: Read slicer definitions from XLSX files.

**GUID**: `{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}`

**File Locations**:
- `xl/slicers/slicer{n}.xml`
- `xl/slicerCaches/slicerCache{n}.xml`

**Files to Create**:
- `openpyxl/worksheet/slicer.py`

**Classes**:
```python
class Slicer(Serialisable):
    tagname = "slicer"
    name = String()
    uid = String()
    caption = String()
    sourceName = String()
    # ... more attributes

class SlicerCache(Serialisable):
    tagname = "slicerCacheDefinition"
    sourceName = String()
    # ... cache items
```

**Note**: Already preserved via Extension class; this adds Python object access.

**Effort**: Medium (200 lines)

---

### B-SLICER-02: Timeline Reader {#b-slicer-02}

**Intent**: Read timeline definitions from XLSX files.

**GUID**: `{7E03D99C-DC04-49d9-9315-930204A7B6E9}`

**File Locations**:
- `xl/timelines/timeline{n}.xml`
- `xl/timelineCaches/timelineCache{n}.xml`

**Classes**:
```python
class Timeline(Serialisable):
    tagname = "timeline"
    name = String()
    uid = String()
    caption = String()
    sourceName = String()
    beginDate = DateTime()
    endDate = DateTime()
```

**Effort**: Medium (150 lines)

---

### B-RICH-01: xlRichValue Reader {#b-rich-01}

**Intent**: Read rich data types (stocks, geography).

**File Locations**:
- `xl/richData/rdrichvalue.xml`
- `xl/richData/rdRichValueTypes.xml`

**Content Type**: `application/vnd.ms-excel.rdRichValue+xml`

**Recommended Approach**: Phase 1 - Binary blob preservation (like VBA).

```python
class RichDataManager:
    def __init__(self):
        self._raw_rdrichvalue_xml = None

    def read(self, archive, path):
        self._raw_rdrichvalue_xml = archive.read(path)

    def write(self, archive):
        if self._raw_rdrichvalue_xml:
            archive.writestr('xl/richData/rdrichvalue.xml',
                           self._raw_rdrichvalue_xml)
```

**Effort**: Low for preservation, High for full parsing

---

### B-SPARK-01: Sparkline Reader {#b-spark-01}

**Intent**: Read sparkline definitions.

**GUID**: `{05C60535-1F16-4FD2-B633-F4F36F0B64E0}`

**Namespace**: `http://schemas.microsoft.com/office/spreadsheetml/2009/9/main` (x14)

**Files to Create**:
- `openpyxl/worksheet/sparkline.py`

**Classes**:
```python
class Sparkline(Serialisable):
    tagname = "sparkline"
    f = String()       # Data range formula
    sqref = String()   # Display cell

class SparklineGroup(Serialisable):
    tagname = "sparklineGroup"
    type = NoneSet(values=['line', 'column', 'stacked'])
    sparklines = Sequence(expected_type=Sparkline)
    # Color attributes...

class SparklineGroupList(Serialisable):
    tagname = "sparklineGroups"
    sparklineGroup = Sequence(expected_type=SparklineGroup)
```

**Note**: Already preserved via Extension class; this adds Python object access.

**Effort**: Medium (150 lines)

---

## P4: Pivot Tables & Tables

### B-PIVOT-01: Pivot Table Creation {#b-pivot-01}

**Intent**: Allow programmatic creation of pivot tables.

**Current State**: Full read/modify support exists (60+ classes). Only creation API missing.

**What's Missing**:
1. Factory/builder methods
2. Automatic ID management
3. Cache generation from source data
4. Table layout calculation
5. Relationship plumbing

**Files to Create**:
- `openpyxl/pivot/builder.py`

**Implementation Approach**:
```python
class PivotTableBuilder:
    def __init__(self, worksheet, source_range, location):
        self.ws = worksheet
        self.source = source_range
        self.location = location

    def add_row_field(self, field_name): ...
    def add_column_field(self, field_name): ...
    def add_value_field(self, field_name, function='sum'): ...
    def build(self) -> TableDefinition: ...
```

**Effort**: High (500+ lines)

---

### B-TABLE-01: Table Creation {#b-table-01}

**Intent**: Allow programmatic creation of tables.

**Current State**: Read/write works. Creation needs better API.

**What's Missing**:
1. `Table.from_headers()` convenience constructor
2. Better write-only mode support
3. Column header auto-extraction

**Files to Modify**:
- `openpyxl/worksheet/table.py`

**Implementation**:
```python
@classmethod
def from_headers(cls, displayName, ref, headers, style=None):
    """Create table with explicit column headers"""
    table = cls(displayName=displayName, ref=ref)
    table._initialise_columns()
    for col, name in zip(table.tableColumns, headers):
        col.name = name
    if style:
        table.tableStyleInfo = TableStyleInfo(name=style)
    return table
```

**Effort**: Low (50 lines)

---

## Implementation Priority

| Priority | Item | Effort | Impact |
|----------|------|--------|--------|
| 1 | B-DYNARR-04 (spill operator) | Low | High - fixes tokenizer crash |
| 2 | B-DYNARR-02 (cell metadata) | Low | High - completes dynamic arrays |
| 3 | B-PRESERVE-03 (Serialisable) | Low | High - foundation for others |
| 4 | B-TABLE-01 (table creation) | Low | Medium - user convenience |
| 5 | B-SPARK-01 (sparklines) | Medium | Medium - popular feature |
| 6 | B-COMMENT-01/02 (threaded) | Medium | Medium - modern Excel |
| 7 | B-PRESERVE-04 (DrawingML) | Medium | High - stops data loss |
| 8 | B-SLICER-01/02 | Medium | Medium - popular feature |
| 9 | B-RICH-01 (rich data) | Low-Med | Low - niche feature |
| 10 | B-CHART-01-06 | High | Medium - new architecture |
| 11 | B-PIVOT-01 | High | High - complex feature |

---

## Constants to Add

In `openpyxl/xml/constants.py`:

```python
# Namespaces
X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XLTHREADED_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
CHARTEX_NS = "http://schemas.microsoft.com/office/drawing/2014/chartex"
XLRD_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
SLICER_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/slicer"
TIMELINE_NS = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/timeline"

# Content Types
PERSON_TYPE = "application/vnd.ms-excel.person+xml"
THREADED_COMMENTS_TYPE = "application/vnd.ms-excel.threadedcomments+xml"
CHARTEX_TYPE = "application/vnd.ms-office.chartex+xml"
RICHVALUE_TYPE = "application/vnd.ms-excel.rdRichValue+xml"
SLICER_TYPE = "application/vnd.ms-excel.slicer+xml"
TIMELINE_TYPE = "application/vnd.ms-excel.Timeline+xml"

# Relationship Types
PERSON_REL = "http://schemas.microsoft.com/office/2017/10/relationships/person"
THREADED_COMMENT_REL = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
```
