# Copyright (c) 2010-2024 openpyxl

"""
Support for extended charts (chartex) in XLSX files.

Extended charts (chartex) include modern chart types introduced in Office 2016:
- Waterfall charts
- Funnel charts
- Treemap charts
- Sunburst charts
- Box & Whisker charts
- Histogram/Pareto charts
- Map charts
- Region maps

These use a different XML namespace and structure than traditional charts:
- Namespace: http://schemas.microsoft.com/office/drawing/2014/chartex
- Content Type: application/vnd.ms-office.chartex+xml
- Root element: <cx:chartSpace>

This module provides binary blob preservation for round-trip fidelity.
Full parsing support is complex and would require 500+ lines of new code.
"""

from copy import deepcopy

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.xml.functions import Element
from openpyxl.xml.constants import CHARTEX_NS, CHARTEX_TYPE, CHARTEX_REL


class ChartExSpace(Serialisable):
    """
    Represents an extended chart (chartex) with raw XML preservation.

    This class preserves the raw XML for extended charts that cannot be
    fully parsed. It allows round-trip fidelity for files containing
    modern chart types like waterfall, funnel, treemap, etc.

    Usage:
        # From XML
        chartex = ChartExSpace.from_tree(node)

        # To XML
        xml = chartex.to_tree()
    """

    tagname = "chartSpace"
    namespace = CHARTEX_NS

    _path = "/xl/charts/chartEx{0}.xml"
    mime_type = CHARTEX_TYPE
    _rel_type = CHARTEX_REL

    _id = None

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
        tag = "{%s}%s" % (self.namespace, tagname or self.tagname)
        return Element(tag)

    @property
    def path(self):
        """Return path within the archive"""
        return self._path.format(self._id)

    def __bool__(self):
        """Return True if this contains chart data"""
        return self._content is not None


def is_chartex(node):
    """Check if an XML node is an extended chart"""
    if node is None:
        return False
    tag = node.tag
    if isinstance(tag, str):
        return CHARTEX_NS in tag or "chartex" in tag.lower()
    return False
