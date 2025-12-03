# Copyright (c) 2010-2024 openpyxl

"""
Support for slicers in XLSX files.

Slicers provide interactive filtering for pivot tables and tables.
They are stored in xl/slicers/slicer{n}.xml files and referenced
via extensions in worksheet XML.
"""

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import String, Integer, Bool, Float
from openpyxl.descriptors.sequence import Sequence


from openpyxl.xml.constants import X14_NS, SLICER_REL, SLICER_CACHE_REL

# Slicer extension GUID
SLICER_GUID = "{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}"
SLICER_LIST_GUID = "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}"


class Slicer(Serialisable):
    """
    Represents a single slicer definition.

    Attributes
    ----------
    name : str
        The unique name of the slicer
    uid : str
        A GUID uniquely identifying this slicer
    cache : str
        Reference to the slicer cache name
    caption : str, optional
        The caption displayed on the slicer
    startItem : int, optional
        Starting item index
    columnCount : int, optional
        Number of columns in the slicer
    showCaption : bool, optional
        Whether to show the caption
    level : int, optional
        For OLAP pivot tables, the level to filter
    style : str, optional
        The slicer style name
    lockedPosition : bool, optional
        Whether position is locked
    rowHeight : int, optional
        Height of each row in EMUs
    """

    tagname = "slicer"
    namespace = X14_NS

    name = String()
    uid = String(allow_none=True)
    cache = String()
    caption = String(allow_none=True)
    startItem = Integer(allow_none=True)
    columnCount = Integer(allow_none=True)
    showCaption = Bool(allow_none=True)
    level = Integer(allow_none=True)
    style = String(allow_none=True)
    lockedPosition = Bool(allow_none=True)
    rowHeight = Integer(allow_none=True)

    __attrs__ = ('name', 'uid', 'cache', 'caption', 'startItem', 'columnCount',
                 'showCaption', 'level', 'style', 'lockedPosition', 'rowHeight')

    def __init__(
        self,
        name=None,
        uid=None,
        cache=None,
        caption=None,
        startItem=None,
        columnCount=None,
        showCaption=None,
        level=None,
        style=None,
        lockedPosition=None,
        rowHeight=None,
    ):
        self.name = name
        self.uid = uid
        self.cache = cache
        self.caption = caption
        self.startItem = startItem
        self.columnCount = columnCount
        self.showCaption = showCaption
        self.level = level
        self.style = style
        self.lockedPosition = lockedPosition
        self.rowHeight = rowHeight


class SlicerList(Serialisable):
    """
    Container for slicers in a worksheet.

    This appears in the worksheet's extension list under the slicer GUID.
    """

    tagname = "slicerList"
    namespace = X14_NS

    slicer = Sequence(expected_type=Slicer)

    __elements__ = ('slicer',)

    _id = None
    _path = "/xl/slicers/slicer{0}.xml"
    mime_type = "application/vnd.ms-excel.slicer+xml"
    _rel_type = SLICER_REL

    def __init__(self, slicer=()):
        self.slicer = slicer

    def __len__(self):
        return len(self.slicer)

    def __bool__(self):
        return bool(self.slicer)

    @property
    def path(self):
        """Return path within the archive"""
        return self._path.format(self._id)


class SlicerCacheItem(Serialisable):
    """
    Represents an item in a slicer cache.
    """

    tagname = "i"

    n = String(allow_none=True)  # Item name/value
    nd = Bool(allow_none=True)   # No data

    __attrs__ = ('n', 'nd')

    def __init__(self, n=None, nd=None):
        self.n = n
        self.nd = nd


class SlicerCacheData(Serialisable):
    """
    Contains the cache data for a slicer.
    """

    tagname = "data"

    # Using simplified structure - real structure is more complex
    # with tabular/OLAP/pivot variations

    def __init__(self):
        pass


class SlicerCacheDefinition(Serialisable):
    """
    Defines the cache for a slicer.

    Slicer caches store the source data references and item lists
    for slicer filtering. They are stored in xl/slicerCaches/.
    """

    tagname = "slicerCacheDefinition"

    name = String()
    uid = String(allow_none=True)
    sourceName = String(allow_none=True)

    _id = None
    _path = "/xl/slicerCaches/slicerCache{0}.xml"
    mime_type = "application/vnd.ms-excel.slicerCache+xml"
    _rel_type = "http://schemas.microsoft.com/office/2007/relationships/slicerCache"

    __attrs__ = ('name', 'uid', 'sourceName')

    def __init__(
        self,
        name=None,
        uid=None,
        sourceName=None,
    ):
        self.name = name
        self.uid = uid
        self.sourceName = sourceName

    @property
    def path(self):
        """Return path within the archive"""
        return self._path.format(self._id)
