# Task: T-PRESERVE-01

## Header

| Field | Value |
|-------|-------|
| **ID** | T-PRESERVE-01 |
| **Parent** | B-PRESERVE-01 |
| **State** | DONE |
| **Created** | 2024-12-03 |
| **Updated** | 2024-12-03 |

## Objective

Stop discarding extension list (extLst) content. Currently openpyxl only stores the URI of extensions, throwing away all content inside `<ext>` tags. This causes sparklines, slicers, timelines, and other modern features to be lost on round-trip.

## Acceptance Criteria

- [x] Extension class preserves raw XML content
- [x] ExtensionList round-trips without data loss
- [x] Warning still issued for unsupported extensions (but data preserved)
- [x] Unit tests for Extension/ExtensionList preservation
- [x] Round-trip test with file containing sparklines

## Context

From `WOTAN/docs/vision.md` - this is the #1 architectural gap.

### Current Behavior

```python
# openpyxl/descriptors/excel.py
class Extension(Serialisable):
    uri = String()  # Only URI stored, CONTENT DISCARDED
```

When reading:
1. Warning: "Sparkline Group extension is not supported and will be removed"
2. Content inside `<ext>` tag thrown away
3. On write: nothing written back

### Target Behavior

```python
class Extension(Serialisable):
    uri = String()
    _content = None  # Raw XML element preserved

    @classmethod
    def from_tree(cls, node):
        obj = cls(uri=node.get('uri'))
        obj._content = copy.deepcopy(node)  # Keep everything
        return obj

    def to_tree(self):
        if self._content is not None:
            return self._content  # Write back unchanged
        # ... fallback for programmatically created extensions
```

### Related Files

- `openpyxl/descriptors/excel.py` - Extension, ExtensionList classes
- `openpyxl/worksheet/_reader.py` - parse_extensions() method
- `openpyxl/xml/constants.py` - EXT_TYPES dict

## Outcome

### Implementation Summary

Modified the Extension class to preserve raw XML content via deepcopy for round-trip fidelity. Key changes:

1. **`openpyxl/descriptors/excel.py`**: Added `_content` field to Extension, implemented `from_tree()` to preserve the full XML element, and `to_tree()` to return preserved content unchanged.

2. **`openpyxl/worksheet/_reader.py`**: Changed warning message from "will be removed" to "will be preserved". Store ExtensionList on parser for binding to worksheet.

3. **`openpyxl/worksheet/_writer.py`**: Added `write_extensions()` method to serialize preserved extensions back to XML.

4. **`openpyxl/worksheet/worksheet.py`**: Added `extensions` attribute to store ExtensionList.

### Files Changed

- `openpyxl/descriptors/excel.py` (+39 lines)
- `openpyxl/descriptors/tests/test_excel.py` (+106 lines)
- `openpyxl/worksheet/_reader.py` (+5/-2 lines)
- `openpyxl/worksheet/_writer.py` (+10 lines)
- `openpyxl/worksheet/worksheet.py` (+1 line)
- `openpyxl/worksheet/tests/test_read_only.py` (+1 line)
- `openpyxl/worksheet/tests/test_reader.py` (+40 lines)
- `openpyxl/worksheet/tests/test_writer.py` (+35 lines)

### Test Results

All 706 tests pass (worksheet + descriptors tests).

