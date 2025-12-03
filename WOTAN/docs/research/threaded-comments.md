# Threaded Comments in XLSX

Research document for implementing threaded comments support in openpyxl.

> **Implementation Status**: ✅ **COMPLETE** (2024-12-03)
> - ThreadedComment/ThreadedCommentList: `openpyxl/comments/threaded.py`
> - Person/PersonList: `openpyxl/comments/person.py`
> - Reader: `openpyxl/reader/excel.py`
> - Writer: `openpyxl/writer/excel.py`

## Overview

Threaded comments (also called "modern comments") were introduced in Excel 365/2019 to replace the legacy comment system. They provide:
- Reply threading (parent-child relationships)
- User identity tracking via persons list
- @mentions support
- Rich text formatting
- Done/resolved status

## Current openpyxl Behavior (with WOTAN)

Threaded comments now survive round-trip. The thread structure, authors (persons), and reply relationships are fully preserved.

## Required XLSX Components

### 1. Content Types (`[Content_Types].xml`)

```xml
<Override PartName="/xl/threadedComments/threadedComment1.xml"
          ContentType="application/vnd.ms-excel.threadedcomments+xml"/>
<Override PartName="/xl/persons/person.xml"
          ContentType="application/vnd.ms-excel.person+xml"/>
```

### 2. Workbook Relationships (`xl/_rels/workbook.xml.rels`)

```xml
<Relationship Id="rId{n}"
              Type="http://schemas.microsoft.com/office/2017/10/relationships/person"
              Target="persons/person.xml"/>
```

### 3. Worksheet Relationships (`xl/worksheets/_rels/sheet{n}.xml.rels`)

```xml
<Relationship Id="rId{n}"
              Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
              Target="../threadedComments/threadedComment{n}.xml"/>
```

### 4. Person List (`xl/persons/person.xml`)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
    <person displayName="John Doe"
            id="{12345678-1234-1234-1234-123456789ABC}"
            userId="user@example.com"
            providerId="PeoplePicker"/>
    <person displayName="Jane Smith"
            id="{87654321-4321-4321-4321-CBA987654321}"
            userId="jane@example.com"
            providerId="PeoplePicker"/>
</personList>
```

### 5. Threaded Comments (`xl/threadedComments/threadedComment{n}.xml`)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
    <threadedComment ref="A1"
                     dT="2024-01-15T10:30:00.000"
                     personId="{12345678-1234-1234-1234-123456789ABC}"
                     id="{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}">
        <text>This is the main comment</text>
    </threadedComment>
    <threadedComment ref="A1"
                     dT="2024-01-15T10:35:00.000"
                     personId="{87654321-4321-4321-4321-CBA987654321}"
                     id="{FFFFFFFF-GGGG-HHHH-IIII-JJJJJJJJJJJJ}"
                     parentId="{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}">
        <text>This is a reply</text>
    </threadedComment>
</ThreadedComments>
```

### 6. Legacy Comment Fallback (`xl/comments{n}.xml`)

Excel maintains a legacy comment for backwards compatibility:

```xml
<comment ref="A1" authorId="0" shapeId="0"
         xr:uid="{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}">
    <text>
        <t>[Threaded comment]

Your version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924

Comment:
    This is the main comment
Reply:
    This is a reply</t>
    </text>
</comment>
```

## XML Schema Details

### CT_PersonList

Container for person definitions.

```xml
<xsd:complexType name="CT_PersonList">
    <xsd:sequence>
        <xsd:element name="person" type="CT_Person" minOccurs="0" maxOccurs="unbounded"/>
        <xsd:element name="extLst" type="x:CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
```

### CT_Person

| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| displayName | string | Yes | Display name shown in UI |
| id | GUID | Yes | Unique identifier (uppercase) |
| userId | string | No | User identifier (email, etc.) |
| providerId | string | No | Identity provider (e.g., "PeoplePicker") |

### CT_ThreadedComments

Container for threaded comments on a worksheet.

```xml
<xsd:complexType name="CT_ThreadedComments">
    <xsd:sequence>
        <xsd:element name="threadedComment" type="CT_ThreadedComment"
                     minOccurs="0" maxOccurs="unbounded"/>
        <xsd:element name="extLst" type="x:CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
</xsd:complexType>
```

### CT_ThreadedComment

| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| ref | string | Yes | Cell reference (e.g., "A1") |
| dT | dateTime | No | UTC timestamp |
| personId | GUID | Yes | Reference to person id |
| id | GUID | Yes | Unique comment identifier (uppercase) |
| parentId | GUID | No | Parent comment id (for replies) |
| done | boolean | No | Resolved/done status |

| Child Element | Type | Description |
|---------------|------|-------------|
| text | string | Comment text content |
| mentions | CT_ThreadedCommentMentions | @mentions in text |
| extLst | CT_ExtensionList | Extensions |

### CT_ThreadedCommentMentions

Container for @mention references.

| Child Element | Type | Description |
|---------------|------|-------------|
| mention | CT_Mention | Individual @mention |

### CT_Mention

| Attribute | Type | Required | Description |
|-----------|------|----------|-------------|
| mentionpersonId | GUID | Yes | Person being mentioned |
| mentionId | GUID | Yes | Unique mention identifier |
| startIndex | uint | Yes | Start position in text |
| length | uint | Yes | Length of mention in text |

## Namespaces

| Prefix | Namespace URI |
|--------|---------------|
| (default) | http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments |
| xr | http://schemas.microsoft.com/office/spreadsheetml/2014/revision |

## Relationship Types

| Type | URI |
|------|-----|
| Person | http://schemas.microsoft.com/office/2017/10/relationships/person |
| ThreadedComment | http://schemas.microsoft.com/office/2017/10/relationships/threadedComment |

## Content Types

| Type | MIME Type |
|------|-----------|
| ThreadedComments | application/vnd.ms-excel.threadedcomments+xml |
| Person | application/vnd.ms-excel.person+xml |

## Implementation Notes

### GUID Format

GUIDs must be **uppercase** and wrapped in curly braces:
- Correct: `{12345678-1234-1234-1234-123456789ABC}`
- Incorrect: `{12345678-1234-1234-1234-123456789abc}`

### Element Ordering

The `legacyDrawing` element must appear before `tableParts` in the worksheet XML.

### VML Drawing

Threaded comments still require VML shapes for the visual comment indicators. The VML drawing must be properly linked via the `legacyDrawing` element in the worksheet.

### Threading Logic

- First comment in a thread: No `parentId` attribute
- Reply comments: `parentId` references the root comment's `id`
- All comments in same thread share the same `ref` (cell reference)

## Implementation Plan for openpyxl

### Phase 1: Preservation (B-COMMENT-01)
1. Create `openpyxl/comments/threaded.py`
2. Parse `xl/threadedComments/threadedComment*.xml`
3. Store threaded comments on worksheet
4. Preserve on round-trip

### Phase 2: Person Support (B-COMMENT-02)
1. Create `openpyxl/comments/person.py`
2. Parse `xl/persons/person.xml`
3. Store persons on workbook
4. Link to threaded comments via personId

### Phase 3: Writer (B-COMMENT-03)
1. Write threadedComment*.xml files
2. Write person.xml
3. Create relationships
4. Maintain legacy comment fallback

## References

- [MS-XLSX: Threaded Comments](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/66e1875d-c60a-48eb-bf88-41066d45fea8)
- [MS-XLSX: CT_ThreadedComments](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/0cfb2f05-87a2-4b5f-a9ad-fb11ca39e2f8)
- [MS-XLSX: Threaded Comments Schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/adb84732-9fc8-48b6-bddc-6b0bcdaad940)
- [OpenXML Adding Threaded Comments](https://stackoverflow.com/questions/71769165/openxml-adding-notes-and-threaded-comments)

## Test Files Needed

Create Excel 365 files demonstrating:
1. Simple threaded comment (single comment)
2. Comment thread with replies
3. Multiple threads on same sheet
4. Comments with @mentions
5. Resolved/done comments
6. Comments with formatting

Store in `openpyxl/tests/data/wotan/threaded-comments/`
