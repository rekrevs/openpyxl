# Copyright (c) 2010-2024 openpyxl

"""
Support for timelines in XLSX files.

Timelines provide date-based filtering for pivot tables.
They are stored in xl/timelines/timeline{n}.xml files and referenced
via extensions in worksheet XML.
"""

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import String, Integer, Bool, DateTime
from openpyxl.descriptors.sequence import Sequence


from openpyxl.xml.constants import X15_NS, TIMELINE_REL, TIMELINE_CACHE_REL

# Timeline extension GUID
TIMELINE_GUID = "{7E03D99C-DC04-49d9-9315-930204A7B6E9}"


class Timeline(Serialisable):
    """
    Represents a single timeline definition.

    Attributes
    ----------
    name : str
        The unique name of the timeline
    uid : str
        A GUID uniquely identifying this timeline
    cache : str
        Reference to the timeline cache name
    caption : str, optional
        The caption displayed on the timeline
    showHeader : bool, optional
        Whether to show the header
    showSelectionLabel : bool, optional
        Whether to show the selection label
    showTimeLevel : bool, optional
        Whether to show the time level selector
    showHorizontalScrollbar : bool, optional
        Whether to show horizontal scrollbar
    level : int, optional
        The time level (0=years, 1=quarters, 2=months, 3=days)
    selectionLevel : int, optional
        Current selection level
    scrollPosition : datetime, optional
        Current scroll position
    style : str, optional
        The timeline style name
    """

    tagname = "timeline"
    namespace = X15_NS

    name = String()
    uid = String(allow_none=True)
    cache = String()
    caption = String(allow_none=True)
    showHeader = Bool(allow_none=True)
    showSelectionLabel = Bool(allow_none=True)
    showTimeLevel = Bool(allow_none=True)
    showHorizontalScrollbar = Bool(allow_none=True)
    level = Integer(allow_none=True)
    selectionLevel = Integer(allow_none=True)
    scrollPosition = DateTime(allow_none=True)
    style = String(allow_none=True)

    __attrs__ = ('name', 'uid', 'cache', 'caption', 'showHeader',
                 'showSelectionLabel', 'showTimeLevel', 'showHorizontalScrollbar',
                 'level', 'selectionLevel', 'scrollPosition', 'style')

    def __init__(
        self,
        name=None,
        uid=None,
        cache=None,
        caption=None,
        showHeader=None,
        showSelectionLabel=None,
        showTimeLevel=None,
        showHorizontalScrollbar=None,
        level=None,
        selectionLevel=None,
        scrollPosition=None,
        style=None,
    ):
        self.name = name
        self.uid = uid
        self.cache = cache
        self.caption = caption
        self.showHeader = showHeader
        self.showSelectionLabel = showSelectionLabel
        self.showTimeLevel = showTimeLevel
        self.showHorizontalScrollbar = showHorizontalScrollbar
        self.level = level
        self.selectionLevel = selectionLevel
        self.scrollPosition = scrollPosition
        self.style = style


class TimelineList(Serialisable):
    """
    Container for timelines in a worksheet.

    This appears in the worksheet's extension list under the timeline GUID.
    """

    tagname = "timelines"
    namespace = X15_NS

    timeline = Sequence(expected_type=Timeline)

    __elements__ = ('timeline',)

    _id = None
    _path = "/xl/timelines/timeline{0}.xml"
    mime_type = "application/vnd.ms-excel.timeline+xml"
    _rel_type = TIMELINE_REL

    def __init__(self, timeline=()):
        self.timeline = timeline

    def __len__(self):
        return len(self.timeline)

    def __bool__(self):
        return bool(self.timeline)

    @property
    def path(self):
        """Return path within the archive"""
        return self._path.format(self._id)


class TimelineCacheDefinition(Serialisable):
    """
    Defines the cache for a timeline.

    Timeline caches store the date field references and state
    for timeline filtering. They are stored in xl/timelineCaches/.
    """

    tagname = "timelineCacheDefinition"
    namespace = X15_NS

    name = String()
    uid = String(allow_none=True)
    sourceName = String(allow_none=True)

    _id = None
    _path = "/xl/timelineCaches/timelineCache{0}.xml"
    mime_type = "application/vnd.ms-excel.timelineCache+xml"
    _rel_type = "http://schemas.microsoft.com/office/2011/relationships/timelineCache"

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
