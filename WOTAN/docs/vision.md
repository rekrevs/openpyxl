# WOTAN Vision

**WOTAN** = Working On The Actual Nuances

## Core Mission

**Read any XLSX file, parse it completely into inspectable/modifiable Python objects, and write it back without losing anything.**

This means:
1. **Complete Parsing**: Every XML element becomes a Python object you can inspect
2. **Full Manipulation**: Modify any aspect of the document programmatically
3. **Lossless Round-Trip**: Write back without losing features you didn't touch
4. **Feature Creation**: Create any XLSX feature from scratch

## Current Reality

### What Works Well (Full Round-Trip)

Openpyxl v3.1.5 achieves excellent fidelity for ECMA-376 (2006) core features:

| Feature | Parse | Modify | Write | Status |
|---------|-------|--------|-------|--------|
| Cell values (all types) | Full | Full | Full | Complete |
| Formulas (standard, array, shared) | Full | Full | Full | Complete |
| Styles (fonts, fills, borders) | Full | Full | Full | Complete |
| Merged cells | Full | Full | Full | Complete |
| Data validation | Full | Full | Full | Complete |
| Conditional formatting | Full | Full | Full | Complete |
| AutoFilter | Full | Full | Full | Complete |
| Hyperlinks | Full | Full | Full | Complete |
| Named ranges | Full | Full | Full | Complete |
| Charts (classic types) | Full | Full | Full | Complete |
| Images | Full | Full | Full | Complete |
| Print settings | Full | Full | Full | Complete |
| Sheet/workbook protection | Full | Full | Full | Complete |

### What's Partially Parsed (Limited Manipulation)

| Feature | Parse | Modify | Write | Issue |
|---------|-------|--------|-------|-------|
| Comments | Text only | Text only | Full | Formatting lost on read |
| Pivot tables | Structure only | Limited | Preserve | Cannot create new |
| Tables | Structure only | Limited | Preserve | Cannot create new |
| Theme | Binary blob | None | Pass-through | Cannot modify colors |
| VBA | Binary archive | None | Pass-through | Cannot create/edit macros |

### What's LOST on Round-Trip (Critical Gaps)

| Feature | Read | Write | Impact |
|---------|------|-------|--------|
| **extLst content** | URI only | Dropped | Sparklines, slicers, timelines LOST |
| **Dynamic arrays** | Not parsed | Not written | Modern formulas break |
| **Threaded comments** | Not parsed | Not written | Comment threads LOST |
| **DrawingML shapes** | Not parsed | Not written | Textboxes, shapes LOST |
| **Rich data types** | Not parsed | Not written | Stocks, geography LOST |
| **Modern charts** | Partial | Partial | Some chart types LOST |
| **Cell metadata** | Not parsed | Not written | Dynamic array flags LOST |
| **Unknown XML** | Silently dropped | N/A | Any new features LOST |

### The extLst Problem

The `<extLst>` (extension list) mechanism is how Excel stores modern features. Currently:

```python
# What openpyxl does:
class Extension(Serialisable):
    uri = String()  # Only URI stored, CONTENT DISCARDED

# What happens:
# 1. Read: Warning "Sparkline Group extension is not supported and will be removed"
# 2. Content inside <ext> tag is thrown away
# 3. Write: Nothing written back
```

**This is the #1 architectural gap** - there's no mechanism to preserve unknown XML.

## Target State

### Level 1: Preserve Unknown (Foundation)

Before we can parse everything, we must stop losing things:

1. **Unknown XML preservation**: Store raw XML for elements we don't understand
2. **extLst content preservation**: Keep extension content even if we can't parse it
3. **Pass-through mode**: Allow features to survive round-trip unchanged

### Level 2: Full Parsing (Excel 365 Features)

| Feature | Target |
|---------|--------|
| Dynamic arrays | Full Python objects, create/modify spilled formulas |
| Threaded comments | Full thread model, create replies, @mentions |
| Modern functions | Recognize all 50+ new functions |
| Modern charts | All chart types as Python objects |
| Sparklines | Full sparkline model, create/modify |
| Slicers & timelines | Full slicer model, create/modify |
| Rich data types | Parse linked data, access fields |

### Level 3: Complete Manipulation

| Capability | Target |
|------------|--------|
| Theme modification | Change theme colors/fonts programmatically |
| Pivot table creation | Create pivot tables from scratch |
| Table creation | Create tables from scratch |
| Comment formatting | Full rich text in comments |

## Success Criteria

1. **Zero-loss round-trip**: `load_workbook(f).save(f2)` loses NOTHING
2. **Full inspection**: Every feature accessible as Python object
3. **Full creation**: Create any feature Excel can create
4. **Excel validation**: Output files open correctly in Excel 365
5. **Backwards compatible**: Existing openpyxl code unchanged

## Architecture Principles

### 1. Preserve What You Don't Understand

```python
# BAD (current): Unknown elements silently dropped
for element in tree:
    if element.tag in known_tags:
        process(element)
    # else: gone forever

# GOOD (target): Unknown elements stored for round-trip
for element in tree:
    if element.tag in known_tags:
        process(element)
    else:
        self._unknown_elements.append(element)
```

### 2. Extension Content Preservation

```python
# BAD (current): Only URI kept
class Extension(Serialisable):
    uri = String()

# GOOD (target): Full content preserved
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

### 3. Parse Incrementally

Add parsing for specific features while preserving unknown:
1. First: Preserve sparklines as raw XML (no loss)
2. Then: Parse sparklines into Python objects (full manipulation)
3. Always: Unknown extensions still preserved

## Priority Order

### Phase 0: Stop the Bleeding
- [ ] Unknown XML preservation infrastructure
- [ ] extLst content preservation
- [ ] Round-trip tests for modern Excel files

### Phase 1: Dynamic Arrays (Critical)
- [ ] Cell metadata (`cm` attribute)
- [ ] `metadata.xml` parsing/writing
- [ ] `futureMetadata XLDAPR`
- [ ] Spilled range operator (`#`)

### Phase 2: Modern Functions
- [ ] Add 50+ functions to FORMULAE
- [ ] Document `_xlfn.` prefix handling
- [ ] Test with real Excel files

### Phase 3: Threaded Comments
- [ ] Parse `threadedComments/*.xml`
- [ ] Parse `persons/person.xml`
- [ ] Full thread model
- [ ] Legacy comment fallback

### Phase 4: Visualization
- [ ] Modern chart types
- [ ] Sparklines (full implementation)
- [ ] DrawingML shapes preservation

### Phase 5: Interactive Features
- [ ] Slicers
- [ ] Timelines
- [ ] Rich data types

### Phase 6: Creation
- [ ] Pivot table creation
- [ ] Table creation
- [ ] Theme modification

## Non-Goals

- **XLSB support** - Binary format is separate
- **XLS support** - Legacy format, use xlrd
- **Formula evaluation** - We preserve, not compute
- **VBA creation** - We preserve, not create
- **Real-time collaboration** - Excel-specific

## Key Resources

- [ECMA-376 5th Edition](https://ecma-international.org/publications-and-standards/standards/ecma-376/)
- [MS-XLSX Extensions](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/)
- [MS-OE376 Implementation Notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/)
- Excel 365 for empirical testing
