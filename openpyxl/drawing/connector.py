# Copyright (c) 2010-2024 openpyxl

from copy import deepcopy

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Bool,
    Integer,
    String,
    Alias,
)
from openpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.xml.functions import Element

from .properties import (
    NonVisualDrawingProps,
    NonVisualDrawingShapeProps,
)
from .geometry import ShapeStyle


class RawShapeElement(Serialisable):
    """
    Preserves unparsed shape XML for round-trip fidelity.

    This class is used when a shape cannot be fully parsed due to
    unsupported features. It stores the raw XML element and can
    serialize it back unchanged.
    """

    tagname = "sp"

    def __init__(self):
        self._content = None

    @classmethod
    def from_tree(cls, node):
        """Store the raw XML element for later serialization"""
        obj = cls()
        obj._content = deepcopy(node)
        return obj

    def to_tree(self, tagname=None, idx=None, namespace=None):
        """Return the preserved raw XML element"""
        if self._content is not None:
            return deepcopy(self._content)
        # Fallback to empty element if no content
        return Element(tagname or self.tagname)

class Connection(Serialisable):

    id = Integer()
    idx = Integer()

    def __init__(self,
                 id=None,
                 idx=None,
                ):
        self.id = id
        self.idx = idx


class ConnectorLocking(Serialisable):

    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 extLst=None,
                ):
        self.extLst = extLst


class NonVisualConnectorProperties(Serialisable):

    cxnSpLocks = Typed(expected_type=ConnectorLocking, allow_none=True)
    stCxn = Typed(expected_type=Connection, allow_none=True)
    endCxn = Typed(expected_type=Connection, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 cxnSpLocks=None,
                 stCxn=None,
                 endCxn=None,
                 extLst=None,
                ):
        self.cxnSpLocks = cxnSpLocks
        self.stCxn = stCxn
        self.endCxn = endCxn
        self.extLst = extLst


class ConnectorNonVisual(Serialisable):

    cNvPr = Typed(expected_type=NonVisualDrawingProps, )
    cNvCxnSpPr = Typed(expected_type=NonVisualConnectorProperties, )

    __elements__ = ("cNvPr", "cNvCxnSpPr",)

    def __init__(self,
                 cNvPr=None,
                 cNvCxnSpPr=None,
                ):
        self.cNvPr = cNvPr
        self.cNvCxnSpPr = cNvCxnSpPr


class ConnectorShape(Serialisable):

    tagname = "cxnSp"

    nvCxnSpPr = Typed(expected_type=ConnectorNonVisual)
    spPr = Typed(expected_type=GraphicalProperties)
    style = Typed(expected_type=ShapeStyle, allow_none=True)
    macro = String(allow_none=True)
    fPublished = Bool(allow_none=True)

    def __init__(self,
                 nvCxnSpPr=None,
                 spPr=None,
                 style=None,
                 macro=None,
                 fPublished=None,
                 ):
        self.nvCxnSpPr = nvCxnSpPr
        self.spPr = spPr
        self.style = style
        self.macro = macro
        self.fPublished = fPublished


class ShapeMeta(Serialisable):

    tagname = "nvSpPr"

    cNvPr = Typed(expected_type=NonVisualDrawingProps)
    cNvSpPr = Typed(expected_type=NonVisualDrawingShapeProps)

    def __init__(self, cNvPr=None, cNvSpPr=None):
        self.cNvPr = cNvPr
        self.cNvSpPr = cNvSpPr


class Shape(Serialisable):

    tagname = "sp"

    macro = String(allow_none=True)
    textlink = String(allow_none=True)
    fPublished = Bool(allow_none=True)
    fLocksText = Bool(allow_none=True)
    nvSpPr = Typed(expected_type=ShapeMeta, allow_none=True)
    meta = Alias("nvSpPr")
    spPr = Typed(expected_type=GraphicalProperties)
    graphicalProperties = Alias("spPr")
    style = Typed(expected_type=ShapeStyle, allow_none=True)
    txBody = Typed(expected_type=RichText, allow_none=True)

    # Flag to indicate this is preserved raw content
    _is_raw = False

    def __init__(self,
                 macro=None,
                 textlink=None,
                 fPublished=None,
                 fLocksText=None,
                 nvSpPr=None,
                 spPr=None,
                 style=None,
                 txBody=None,
                ):
        self.macro = macro
        self.textlink = textlink
        self.fPublished = fPublished
        self.fLocksText = fLocksText
        self.nvSpPr = nvSpPr
        self.spPr = spPr
        self.style = style
        self.txBody = txBody
        self._raw_content = None

    @classmethod
    def from_tree(cls, node):
        """
        Parse shape from XML. If parsing fails due to unsupported features,
        fall back to preserving raw XML for round-trip fidelity.
        """
        try:
            return super(Shape, cls).from_tree(node)
        except (TypeError, KeyError, AttributeError):
            # Parsing failed - preserve as raw content
            obj = cls.__new__(cls)
            obj._raw_content = deepcopy(node)
            obj._is_raw = True
            # Initialize all attributes to None/defaults to avoid AttributeError
            obj.macro = None
            obj.textlink = None
            obj.fPublished = None
            obj.fLocksText = None
            obj.nvSpPr = None
            obj.spPr = None
            obj.style = None
            obj.txBody = None
            return obj

    def to_tree(self, tagname=None, idx=None, namespace=None):
        """
        Serialize shape to XML. If this is preserved raw content,
        return the original XML unchanged.
        """
        if getattr(self, '_raw_content', None) is not None:
            return deepcopy(self._raw_content)
        return super().to_tree(tagname, idx, namespace)
