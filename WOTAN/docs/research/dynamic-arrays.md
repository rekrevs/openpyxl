# Dynamic Arrays in XLSX

Research document for implementing dynamic array support in openpyxl.

> **Implementation Status**: ✅ **COMPLETE** (2024-12-03)
> - Cell `cm` attribute: `openpyxl/cell/_writer.py`, `openpyxl/worksheet/_reader.py`
> - metadata.xml: `openpyxl/packaging/metadata.py`
> - Spilled range `#` operator: `openpyxl/formula/tokenizer.py`

## Overview

Dynamic arrays (introduced in Excel 365/2019) allow a single formula to return multiple values that "spill" into adjacent cells. This requires specific XLSX structures that openpyxl now fully supports via WOTAN extensions.

## Current openpyxl Behavior (with WOTAN)

Dynamic arrays now survive round-trip. The cell metadata (`cm` attribute) and metadata.xml are properly preserved.

## Required XLSX Components

### 1. Content Types (`[Content_Types].xml`)

Add metadata content type:
```xml
<Override PartName="/xl/metadata.xml"
          ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
```

### 2. Workbook Relationships (`xl/_rels/workbook.xml.rels`)

Add relationship to metadata.xml:
```xml
<Relationship Id="rId{n}"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
              Target="metadata.xml"/>
```

### 3. Metadata File (`xl/metadata.xml`)

Complete structure for dynamic arrays:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">

    <!-- Metadata type declaration -->
    <metadataTypes count="1">
        <metadataType name="XLDAPR"
                      minSupportedVersion="120000"
                      copy="1"
                      pasteAll="1"
                      pasteValues="1"
                      merge="1"
                      splitFirst="1"
                      rowColShift="1"
                      clearFormats="1"
                      clearComments="1"
                      assign="1"
                      coerce="1"/>
    </metadataTypes>

    <!-- Future metadata for dynamic array properties -->
    <futureMetadata name="XLDAPR" count="1">
        <bk>
            <extLst>
                <ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">
                    <xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>
                </ext>
            </extLst>
        </bk>
    </futureMetadata>

    <!-- Cell metadata index -->
    <cellMetadata count="1">
        <bk>
            <rc t="1" v="0"/>
        </bk>
    </cellMetadata>
</metadata>
```

### 4. Cell Element (`xl/worksheets/sheet{n}.xml`)

Cells with dynamic array formulas need the `cm` attribute:

```xml
<c r="B1" cm="1">
    <f t="array" ref="B1:B10">_xlfn.UNIQUE(A1:A10)</f>
    <v>first_value</v>
</c>
```

**Key attributes:**
- `cm="1"` - Cell Metadata Index, references the cellMetadata block (1-based index)
- `t="array"` - Formula type indicating array formula
- `ref="B1:B10"` - The spill range (anchor cell to last spilled cell)

## XML Schema Details

### metadata Element (CT_Metadata)

Root element containing all metadata structures.

| Child Element | Type | Occurrences | Description |
|---------------|------|-------------|-------------|
| metadataTypes | CT_MetadataTypes | 0-1 | Metadata type definitions |
| metadataStrings | CT_MetadataStrings | 0-1 | String storage |
| mdxMetadata | CT_MdxMetadata | 0-1 | MDX metadata |
| futureMetadata | CT_FutureMetadata | 0-* | Future metadata blocks |
| cellMetadata | CT_MetadataBlocks | 0-1 | Cell metadata records |
| valueMetadata | CT_MetadataBlocks | 0-1 | Value metadata records |
| extLst | CT_ExtensionList | 0-1 | Extensions |

### metadataType Element (CT_MetadataType)

Defines a metadata type and its behavior flags.

| Attribute | Type | Description |
|-----------|------|-------------|
| name | ST_Xstring | Type name (e.g., "XLDAPR") |
| minSupportedVersion | unsignedInt | Minimum Excel version |
| copy | boolean | Copy behavior |
| pasteAll | boolean | Paste all behavior |
| pasteValues | boolean | Paste values behavior |
| merge | boolean | Merge behavior |
| splitFirst | boolean | Split first behavior |
| rowColShift | boolean | Row/column shift behavior |
| clearFormats | boolean | Clear formats behavior |
| clearComments | boolean | Clear comments behavior |
| assign | boolean | Assignment behavior |
| coerce | boolean | Coercion behavior |

### futureMetadata Element (CT_FutureMetadata)

Container for future/extension metadata.

| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| name | ST_Xstring | Yes | Links to metadataType name |
| count | unsignedInt | No | Block count (default 0) |

| Child Element | Type | Occurrences |
|---------------|------|-------------|
| bk | CT_FutureMetadataBlock | 0-* |
| extLst | CT_ExtensionList | 0-1 |

### cellMetadata / valueMetadata (CT_MetadataBlocks)

| Attribute | Type | Description |
|-----------|------|-------------|
| count | unsignedInt | Number of blocks |

Contains `bk` elements, each with `rc` (record) elements:

```xml
<cellMetadata count="1">
    <bk>
        <rc t="1" v="0"/>
    </bk>
</cellMetadata>
```

- `t` (type): 1-based index into metadataTypes
- `v` (value): 0-based index into matching futureMetadata

### DynamicArrayProperties (xda:dynamicArrayProperties)

Part of the `xda` namespace (Office 2019+).

| Attribute | Type | Description |
|-----------|------|-------------|
| fDynamic | boolean | Formula is dynamic (spills) |
| fCollapsed | boolean | Spill range is collapsed |

**Namespace:** `http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray`

**Extension URI:** `{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}`

## Cell cm Attribute

The `cm` attribute on cell elements (`<c>`) is a zero-based index into the cellMetadata blocks:

- `cm="1"` references the first cellMetadata block
- The cellMetadata block's `rc` element links to metadataType and futureMetadata

**Index Chain:**
```
Cell cm="1"
  -> cellMetadata/bk[0]/rc[@t="1", @v="0"]
     -> metadataTypes/metadataType[0] (t=1 is 1-based)
     -> futureMetadata[@name="XLDAPR"]/bk[0] (v=0 is 0-based)
```

## Spilled Range Operator (#)

The `#` operator in formulas refers to a spilled range:

```
=A1#    -- References the entire spill range starting at A1
=SUM(B1#)  -- Sum of all values in B1's spill range
```

The tokenizer should recognize `#` as a postfix operator on cell references.

## Implementation Plan for openpyxl

### Phase 1: Preservation (B-DYNARR-01)
1. Create `openpyxl/packaging/metadata.py`
2. Add `METADATA_NS` constant
3. Add content type to manifest
4. Add workbook relationship
5. Parse and preserve existing metadata.xml on round-trip

### Phase 2: Cell Metadata (B-DYNARR-02)
1. Add `cm` attribute to cell reader/writer
2. Add `metadata_index` property to Cell class
3. Link to metadata infrastructure

### Phase 3: Dynamic Array Properties (B-DYNARR-03)
1. Implement `FutureMetadata` class
2. Implement `DynamicArrayProperties` class
3. Handle XLDAPR namespace

### Phase 4: Formula Support (B-DYNARR-04)
1. Update tokenizer for `#` operator
2. Preserve spilled range references

## References

- [MS-XLSX Metadata Specification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/3dd44d53-847b-402f-a8c7-41a85024caf7)
- [ECMA-376 Part 4 - futureMetadata](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_futureMetadata_topic_ID0E4LU6.html)
- [DynamicArrayProperties Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2019.excel.dynamicarray.dynamicarrayproperties)
- [CellMetaIndex Property](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.celltype.cellmetaindex)
- [Dynamic Array Formulas - Microsoft Support](https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)

## Test Files Needed

Create Excel 365 files demonstrating:
1. Simple UNIQUE formula with spill
2. FILTER formula with multiple columns
3. SORT/SORTBY formulas
4. Nested dynamic array formulas
5. Spilled range operator (`#`) usage
6. Multiple dynamic arrays in same sheet

Store in `openpyxl/tests/data/wotan/dynamic-arrays/`
