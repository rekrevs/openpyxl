# XLSX Format Versions and Compatibility

Research on XLSX format versioning and how to handle it.

## Executive Summary

**Key Finding**: XLSX versioning is a known but manageable issue. The main concerns are:

1. **Transitional vs Strict** - Two conformance classes with different namespaces
2. **ECMA-376 Editions** - 5 editions (2006-2016) with minor differences
3. **Microsoft Extensions** - MS-XLSX adds features beyond the standard
4. **Function Prefixes** - `_xlfn.` and `_xlpm.` for newer Excel functions

**Recommendation**: Focus on Transitional (what everyone uses) and MS-XLSX extensions (what Excel 365 produces). Strict format is rarely used in practice.

---

## Format Hierarchy

```
ECMA-376 (2006)
    ↓
ISO/IEC 29500:2008
    ├── Part 1: Strict variant
    └── Part 4: Transitional variant (backwards compatible)
    ↓
ISO/IEC 29500:2012, 2016 (corrections)
    ↓
MS-XLSX Extensions (continuously updated)
    └── Excel 365 features (dynamic arrays, etc.)
```

---

## Transitional vs Strict

### Transitional (What Everyone Uses)

- **Namespace**: `http://schemas.openxmlformats.org/spreadsheetml/2006/main`
- **Default in**: All versions of Excel, LibreOffice, Google Sheets
- **Purpose**: Backwards compatibility with Excel 2007+ files
- **Dates**: Uses legacy serial date numbers (1900 or 1904 date system)

### Strict (Rarely Used)

- **Namespace**: `http://purl.oclc.org/ooxml/spreadsheetml/main`
- **Available in**: Excel 2013+ (must explicitly save as "Strict")
- **Purpose**: Standards compliance, no legacy baggage
- **Dates**: ISO 8601 format allowed

### Practical Impact

| Aspect | Transitional | Strict |
|--------|--------------|--------|
| Excel default save | Yes | No |
| LibreOffice write | Yes | No |
| Google Sheets | Yes | No |
| openpyxl support | Yes | Limited |
| EPPlus support | Yes | No |

**Conclusion**: Target Transitional. Strict is a non-issue for most users.

### Openpyxl's Current Strict Handling

From [Stack Overflow](https://stackoverflow.com/questions/62800822/openpyxl-cannot-read-strict-open-xml-spreadsheet-format-userwarning-file-conta):

> openpyxl doesn't properly support Strict Open XML Spreadsheet format. When opening files saved in this format, openpyxl doesn't issue an error, but rather prints a warning: "File contains an invalid specification for Sheet1. This will be removed."

**Workaround**: Save file as regular xlsx instead of "Strict Open XML Spreadsheet".

---

## ECMA-376 Editions

| Edition | Date | Key Changes |
|---------|------|-------------|
| 1st | Dec 2006 | Original specification |
| 2nd | Dec 2008 | Aligned with ISO 29500:2008 |
| 3rd | Jun 2011 | Corrections |
| 4th | Dec 2012 | Date handling clarification |
| 5th | Dec 2016 | Minor updates |

### Key Finding

From [Library of Congress](https://www.loc.gov/preservation/digital/formats/fdd/fdd000398.shtml):

> "The specification has had very few changes other than clarifications and corrections to match actual usage in documents since SpreadsheetML was first standardized."

**The main change**: Date handling amendment in 2012 disallowed ISO 8601 dates in Transitional files (they caused compatibility problems).

**Conclusion**: Edition differences are minor. Not a major concern.

---

## Microsoft Extensions (MS-XLSX)

This is where modern Excel features live. The [MS-XLSX specification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/) documents:

- Dynamic arrays
- Modern functions (LAMBDA, LET, etc.)
- Threaded comments
- Rich data types
- New chart types

### Current Version

MS-XLSX v29.0 (September 2025) - continuously updated by Microsoft.

### Extension Namespaces

```
x14:  http://schemas.microsoft.com/office/spreadsheetml/2009/9/main
x15:  http://schemas.microsoft.com/office/spreadsheetml/2010/11/main
xda:  http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray
xr:   http://schemas.microsoft.com/office/spreadsheetml/2017/richdata
```

**Conclusion**: MS-XLSX extensions are the real versioning challenge, not ECMA-376 editions.

---

## Function Prefixes

From [Stack Overflow](https://stackoverflow.com/questions/64294341/what-are-xlfn-and-xlpm-in-excel-mean) and [openpyxl-users](https://groups.google.com/g/openpyxl-users/c/O746AjGV9EY):

### `_xlfn.` Prefix

Used for functions not in original specification. Required for:
- XLOOKUP, XMATCH
- FILTER, SORT, UNIQUE, SEQUENCE
- CONCAT, TEXTJOIN
- IFS, SWITCH, MAXIFS, MINIFS
- And many more

**Example**: `=_xlfn.XLOOKUP(A1,B:B,C:C)`

### `_xlpm.` Prefix

Used for LAMBDA and LET parameter names.

**Example**: `=_xlfn.LET(_xlpm.x,A1,_xlpm.x*2)`

### `_xlfn._xlws.` Prefix

Used for worksheet-scoped functions.

### Openpyxl Handling

Openpyxl correctly preserves these prefixes when reading/writing. The issue is that the FORMULAE validation list doesn't include modern functions, so validation may warn incorrectly.

---

## How Other Libraries Handle This

### XlsxWriter (Python, write-only)

- Writes Transitional format only
- Supports dynamic arrays (added 2020)
- Documents formula prefixes clearly
- No Strict support

### SheetJS (JavaScript)

From their docs:
> "Excel does not follow the specification, and there are additional documents discussing how Excel deviates from the specification."

- Reads both Transitional and Strict
- Writes Transitional
- Handles date format differences

### EPPlus (.NET)

From [EPPlus docs](https://www.epplussoftware.com/):
> "EPPlus does not support the xlsx Strict format."

- Transitional only
- Full threaded comments support
- Dynamic arrays support

### Apache POI (Java)

From [Stack Overflow](https://stackoverflow.com/questions/16836964/does-apache-poi-support-the-strict-variant-of-iso-iec-29500):
> "There's not that much different between the two, but where there is such a difference that matters, POI will have code to munge the low level XML."

- Supports both Transitional and Strict (with munging)
- Most mature OOXML implementation

---

## Community Discussions Summary

### openpyxl-users Google Group

Key threads:

1. **[Office 365 Excel dynamic arrays support?](https://groups.google.com/g/openpyxl-users/c/guBoEY28XWk)**
   - Users asking for dynamic array support
   - Maintainer response: "any particular version of MS Excel is not a target for the library"
   - Known issue: formulas get wrapped in curly braces, breaking workbooks

2. **[On OpenPyXL and Dynamic Arrays!](https://groups.google.com/g/openpyxl-users/c/aacD5eRiP7w)**
   - User investigated XML differences
   - Found: need `metadata.xml`, cell `cm` attribute, XLDAPR namespace
   - Bi-directional spill requires `calcChain` (complex)

3. **[Excel LET function not working](https://groups.google.com/g/openpyxl-users/c/O746AjGV9EY)**
   - Solution: use both `_xlfn.` and `_xlpm.` prefixes
   - Example: `_xlfn.LET(_xlpm.MYVAL,"test",_xlpm.MYVAL)`

4. **[Openpyxl, color and Office 365](https://groups.google.com/g/openpyxl-users/c/kdMFI8Paw6I)**
   - Color formatting issues with Office 365 online
   - Maintainer: "openpyxl's behaviour hasn't changed... the ball is in Microsoft's court"

### Maintainer Philosophy

From various threads:
- Target is ECMA-376/ISO 29500, not specific Excel versions
- Features can be added but require development effort
- Backward compatibility is important
- Project maintained by volunteers

---

## Recommendations for WOTAN

### 1. Format Support Strategy

```
Priority 1: Transitional (current focus) - what everyone uses
Priority 2: MS-XLSX extensions - modern Excel features
Priority 3: Strict format - add basic support/conversion
```

### 2. Version Detection

Add capability to detect format variant:
```python
def detect_xlsx_variant(archive):
    """Detect if XLSX is Transitional or Strict based on namespace."""
    # Check workbook.xml namespace
    # Transitional: schemas.openxmlformats.org
    # Strict: purl.oclc.org/ooxml
```

### 3. Namespace Handling

Add all MS-XLSX extension namespaces to `xml/constants.py`:
```python
# Microsoft extensions
X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
X15_NS = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
XLDAPR_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"
# etc.
```

### 4. Function Prefix Handling

- Document `_xlfn.` and `_xlpm.` usage
- Update FORMULAE list with modern functions
- Preserve prefixes in round-trip (already works)

### 5. Documentation

Document which features require which format version:
| Feature | Minimum Version | Namespace |
|---------|-----------------|-----------|
| Basic XLSX | ECMA-376 Ed 1 | main |
| Sparklines | Excel 2010 | x14 |
| Dynamic Arrays | Excel 365 | xda |
| Threaded Comments | Excel 365 | xr |

---

## Sources

- [XLSX Transitional Format - Library of Congress](https://www.loc.gov/preservation/digital/formats/fdd/fdd000398.shtml)
- [XLSX Strict Format - Library of Congress](https://www.loc.gov/preservation/digital/formats/fdd/fdd000401.shtml)
- [MS-XLSX Specification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/)
- [MS-OE376 Implementation Notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/)
- [openpyxl-users: Dynamic Arrays](https://groups.google.com/g/openpyxl-users/c/aacD5eRiP7w)
- [openpyxl-users: Office 365 Support](https://groups.google.com/g/openpyxl-users/c/guBoEY28XWk)
- [openpyxl-users: LET Function](https://groups.google.com/g/openpyxl-users/c/O746AjGV9EY)
- [Stack Overflow: _xlfn and _xlpm meaning](https://stackoverflow.com/questions/64294341/what-are-xlfn-and-xlpm-in-excel-mean)
- [Stack Overflow: openpyxl Strict format](https://stackoverflow.com/questions/62800822/openpyxl-cannot-read-strict-open-xml-spreadsheet-format)
