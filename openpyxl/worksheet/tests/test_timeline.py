# Copyright (c) 2010-2024 openpyxl

import pytest
from datetime import datetime

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def Timeline():
    from ..timeline import Timeline
    return Timeline


@pytest.fixture
def TimelineList():
    from ..timeline import TimelineList
    return TimelineList


@pytest.fixture
def TimelineCacheDefinition():
    from ..timeline import TimelineCacheDefinition
    return TimelineCacheDefinition


class TestTimeline:

    def test_ctor(self, Timeline):
        timeline = Timeline(
            name="Timeline_Date",
            cache="NativeTimeline_Date",
            caption="Date",
            showHeader=True,
            showSelectionLabel=True,
            showTimeLevel=True,
            level=2,
            style="TimeSlicerStyleLight1"
        )
        assert timeline.name == "Timeline_Date"
        assert timeline.cache == "NativeTimeline_Date"
        assert timeline.caption == "Date"
        assert timeline.showHeader is True
        assert timeline.showSelectionLabel is True
        assert timeline.showTimeLevel is True
        assert timeline.level == 2
        assert timeline.style == "TimeSlicerStyleLight1"

    def test_from_xml(self, Timeline):
        src = """
        <timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                name="Timeline_OrderDate"
                uid="{12345678-1234-1234-1234-123456789012}"
                cache="NativeTimeline_OrderDate"
                caption="Order Date"
                showHeader="1"
                showSelectionLabel="1"
                showTimeLevel="1"
                level="2"
                style="TimeSlicerStyleDark1"/>
        """
        node = fromstring(src)
        timeline = Timeline.from_tree(node)
        assert timeline.name == "Timeline_OrderDate"
        assert timeline.uid == "{12345678-1234-1234-1234-123456789012}"
        assert timeline.cache == "NativeTimeline_OrderDate"
        assert timeline.caption == "Order Date"
        assert timeline.showHeader is True
        assert timeline.showSelectionLabel is True
        assert timeline.showTimeLevel is True
        assert timeline.level == 2
        assert timeline.style == "TimeSlicerStyleDark1"

    def test_to_xml(self, Timeline):
        timeline = Timeline(
            name="Timeline_ShipDate",
            cache="NativeTimeline_ShipDate",
            caption="Ship Date"
        )
        tree = timeline.to_tree()
        assert tree is not None
        assert tree.get("name") == "Timeline_ShipDate"


class TestTimelineList:

    def test_ctor(self, TimelineList, Timeline):
        timelines = TimelineList(timeline=[
            Timeline(name="Timeline_1", cache="Cache_1"),
            Timeline(name="Timeline_2", cache="Cache_2"),
        ])
        assert len(timelines) == 2

    def test_from_xml(self, TimelineList):
        src = """
        <timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
            <timeline name="Timeline_A" cache="Cache_A" caption="Column A"/>
            <timeline name="Timeline_B" cache="Cache_B" caption="Column B"/>
        </timelines>
        """
        node = fromstring(src)
        timeline_list = TimelineList.from_tree(node)
        assert len(timeline_list) == 2
        assert timeline_list.timeline[0].name == "Timeline_A"
        assert timeline_list.timeline[1].name == "Timeline_B"

    def test_bool(self, TimelineList, Timeline):
        empty = TimelineList()
        assert not empty

        with_timelines = TimelineList(timeline=[
            Timeline(name="Test", cache="Test")
        ])
        assert with_timelines

    def test_path(self, TimelineList):
        tl = TimelineList()
        tl._id = 1
        assert tl.path == "/xl/timelines/timeline1.xml"


class TestTimelineCacheDefinition:

    def test_ctor(self, TimelineCacheDefinition):
        cache = TimelineCacheDefinition(
            name="NativeTimeline_Date",
            sourceName="Date"
        )
        assert cache.name == "NativeTimeline_Date"
        assert cache.sourceName == "Date"

    def test_path(self, TimelineCacheDefinition):
        cache = TimelineCacheDefinition(name="Test")
        cache._id = 1
        assert cache.path == "/xl/timelineCaches/timelineCache1.xml"

    def test_from_xml(self, TimelineCacheDefinition):
        src = """
        <timelineCacheDefinition
            name="NativeTimeline_OrderDate"
            sourceName="Order Date"/>
        """
        node = fromstring(src)
        cache = TimelineCacheDefinition.from_tree(node)
        assert cache.name == "NativeTimeline_OrderDate"
        assert cache.sourceName == "Order Date"
