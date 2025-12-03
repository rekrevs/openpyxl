# Copyright (c) 2010-2024 openpyxl

"""
Metadata handling for XLSX files.

Implements reading and writing of xl/metadata.xml which is required
for dynamic arrays and other modern Excel features.
"""

from openpyxl.descriptors import (
    Integer,
    Bool,
    String,
    Sequence,
)
from openpyxl.descriptors.base import Typed
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import fromstring


# Relationship type for metadata
METADATA_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
# Content type for metadata
METADATA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
# Namespace for dynamic arrays (Office 2019+)
DYNAMIC_ARRAY_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"
# Extension URI for XLDAPR (dynamic array properties)
XLDAPR_URI = "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"


class MetadataType(Serialisable):
    """
    Defines a metadata type and its behavior flags.

    Per ECMA-376 18.9.17 CT_MetadataType
    """

    tagname = "metadataType"

    name = String()
    minSupportedVersion = Integer(allow_none=True)
    ghostRow = Bool(allow_none=True)
    ghostCol = Bool(allow_none=True)
    edit = Bool(allow_none=True)
    delete = Bool(allow_none=True)
    copy = Bool(allow_none=True)
    pasteAll = Bool(allow_none=True)
    pasteFormulas = Bool(allow_none=True)
    pasteValues = Bool(allow_none=True)
    pasteFormats = Bool(allow_none=True)
    pasteComments = Bool(allow_none=True)
    pasteDataValidation = Bool(allow_none=True)
    pasteBorders = Bool(allow_none=True)
    pasteColWidths = Bool(allow_none=True)
    pasteNumberFormats = Bool(allow_none=True)
    merge = Bool(allow_none=True)
    splitFirst = Bool(allow_none=True)
    splitAll = Bool(allow_none=True)
    rowColShift = Bool(allow_none=True)
    clearAll = Bool(allow_none=True)
    clearFormats = Bool(allow_none=True)
    clearContents = Bool(allow_none=True)
    clearComments = Bool(allow_none=True)
    assign = Bool(allow_none=True)
    coerce = Bool(allow_none=True)
    adjust = Bool(allow_none=True)
    cellMeta = Bool(allow_none=True)

    def __init__(self,
                 name=None,
                 minSupportedVersion=None,
                 ghostRow=None,
                 ghostCol=None,
                 edit=None,
                 delete=None,
                 copy=None,
                 pasteAll=None,
                 pasteFormulas=None,
                 pasteValues=None,
                 pasteFormats=None,
                 pasteComments=None,
                 pasteDataValidation=None,
                 pasteBorders=None,
                 pasteColWidths=None,
                 pasteNumberFormats=None,
                 merge=None,
                 splitFirst=None,
                 splitAll=None,
                 rowColShift=None,
                 clearAll=None,
                 clearFormats=None,
                 clearContents=None,
                 clearComments=None,
                 assign=None,
                 coerce=None,
                 adjust=None,
                 cellMeta=None,
                 ):
        self.name = name
        self.minSupportedVersion = minSupportedVersion
        self.ghostRow = ghostRow
        self.ghostCol = ghostCol
        self.edit = edit
        self.delete = delete
        self.copy = copy
        self.pasteAll = pasteAll
        self.pasteFormulas = pasteFormulas
        self.pasteValues = pasteValues
        self.pasteFormats = pasteFormats
        self.pasteComments = pasteComments
        self.pasteDataValidation = pasteDataValidation
        self.pasteBorders = pasteBorders
        self.pasteColWidths = pasteColWidths
        self.pasteNumberFormats = pasteNumberFormats
        self.merge = merge
        self.splitFirst = splitFirst
        self.splitAll = splitAll
        self.rowColShift = rowColShift
        self.clearAll = clearAll
        self.clearFormats = clearFormats
        self.clearContents = clearContents
        self.clearComments = clearComments
        self.assign = assign
        self.coerce = coerce
        self.adjust = adjust
        self.cellMeta = cellMeta


class MetadataTypes(Serialisable):
    """
    Container for metadata type definitions.

    Per ECMA-376 18.9.18 CT_MetadataTypes
    """

    tagname = "metadataTypes"

    count = Integer(allow_none=True)
    metadataType = Sequence(expected_type=MetadataType)

    __elements__ = ('metadataType',)

    def __init__(self, count=None, metadataType=()):
        self.count = count
        self.metadataType = metadataType


class MetadataRecord(Serialisable):
    """
    A metadata record linking to type and value indices.

    Per ECMA-376 18.9.14 CT_MetadataRecord
    """

    tagname = "rc"

    t = Integer()  # 1-based index into metadataTypes
    v = Integer()  # 0-based index into futureMetadata

    def __init__(self, t=None, v=None):
        self.t = t
        self.v = v


class MetadataBlock(Serialisable):
    """
    A block of metadata records.

    Per ECMA-376 18.9.3 CT_MetadataBlock
    """

    tagname = "bk"

    rc = Sequence(expected_type=MetadataRecord)

    __elements__ = ('rc',)

    def __init__(self, rc=()):
        self.rc = rc


class MetadataBlocks(Serialisable):
    """
    Container for metadata blocks (used for cellMetadata and valueMetadata).

    Per ECMA-376 18.9.4 CT_MetadataBlocks
    """

    count = Integer(allow_none=True)
    bk = Sequence(expected_type=MetadataBlock)

    __elements__ = ('bk',)

    def __init__(self, count=None, bk=()):
        self.count = count
        self.bk = bk


class CellMetadata(MetadataBlocks):
    """Cell metadata container."""
    tagname = "cellMetadata"


class ValueMetadata(MetadataBlocks):
    """Value metadata container."""
    tagname = "valueMetadata"


class FutureMetadataBlock(Serialisable):
    """
    A block of future metadata.

    Per ECMA-376 18.9.10 CT_FutureMetadataBlock
    """

    tagname = "bk"

    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('extLst',)

    def __init__(self, extLst=None):
        self.extLst = extLst


class FutureMetadata(Serialisable):
    """
    Container for future metadata (extension point).

    Per ECMA-376 18.9.9 CT_FutureMetadata
    """

    tagname = "futureMetadata"

    name = String()
    count = Integer(allow_none=True)
    bk = Sequence(expected_type=FutureMetadataBlock)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('bk', 'extLst')

    def __init__(self, name=None, count=None, bk=(), extLst=None):
        self.name = name
        self.count = count
        self.bk = bk
        self.extLst = extLst


class Metadata(Serialisable):
    """
    Root element for metadata.xml.

    Per ECMA-376 18.9.8 CT_Metadata
    """

    tagname = "metadata"

    metadataTypes = Typed(expected_type=MetadataTypes, allow_none=True)
    metadataStrings = None  # Not commonly used, preserve raw if present
    mdxMetadata = None  # MDX metadata, preserve raw if present
    futureMetadata = Sequence(expected_type=FutureMetadata)
    cellMetadata = Typed(expected_type=CellMetadata, allow_none=True)
    valueMetadata = Typed(expected_type=ValueMetadata, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('metadataTypes', 'futureMetadata', 'cellMetadata',
                    'valueMetadata', 'extLst')

    def __init__(self,
                 metadataTypes=None,
                 metadataStrings=None,
                 mdxMetadata=None,
                 futureMetadata=(),
                 cellMetadata=None,
                 valueMetadata=None,
                 extLst=None,
                 ):
        self.metadataTypes = metadataTypes
        self.metadataStrings = metadataStrings
        self.mdxMetadata = mdxMetadata
        self.futureMetadata = futureMetadata
        self.cellMetadata = cellMetadata
        self.valueMetadata = valueMetadata
        self.extLst = extLst

    # For use with relationship system
    mime_type = METADATA_TYPE
    rel_type = METADATA_REL
    _path = "/xl/metadata.xml"

    @property
    def path(self):
        return self._path


def read_metadata(archive):
    """
    Read metadata.xml from archive if present.

    Args:
        archive: ZipFile archive

    Returns:
        Metadata object or None if not present
    """
    try:
        src = archive.read("xl/metadata.xml")
        node = fromstring(src)
        return Metadata.from_tree(node)
    except KeyError:
        return None
