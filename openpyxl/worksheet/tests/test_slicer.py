# Copyright (c) 2010-2024 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def Slicer():
    from ..slicer import Slicer
    return Slicer


@pytest.fixture
def SlicerList():
    from ..slicer import SlicerList
    return SlicerList


@pytest.fixture
def SlicerCacheDefinition():
    from ..slicer import SlicerCacheDefinition
    return SlicerCacheDefinition


class TestSlicer:

    def test_ctor(self, Slicer):
        slicer = Slicer(
            name="Slicer_Category",
            cache="Slicer_Category",
            caption="Category",
            columnCount=2,
            showCaption=True,
            style="SlicerStyleLight1"
        )
        assert slicer.name == "Slicer_Category"
        assert slicer.cache == "Slicer_Category"
        assert slicer.caption == "Category"
        assert slicer.columnCount == 2
        assert slicer.showCaption is True
        assert slicer.style == "SlicerStyleLight1"

    def test_from_xml(self, Slicer):
        src = """
        <slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
                name="Slicer_Region"
                cache="Slicer_Region"
                caption="Region"
                columnCount="1"
                showCaption="1"
                style="SlicerStyleDark2"/>
        """
        node = fromstring(src)
        slicer = Slicer.from_tree(node)
        assert slicer.name == "Slicer_Region"
        assert slicer.cache == "Slicer_Region"
        assert slicer.caption == "Region"
        assert slicer.columnCount == 1
        assert slicer.showCaption is True
        assert slicer.style == "SlicerStyleDark2"

    def test_to_xml(self, Slicer):
        slicer = Slicer(
            name="Slicer_Product",
            cache="Slicer_Product",
            caption="Product"
        )
        tree = slicer.to_tree()
        assert tree is not None
        assert tree.get("name") == "Slicer_Product"


class TestSlicerList:

    def test_ctor(self, SlicerList, Slicer):
        slicers = SlicerList(slicer=[
            Slicer(name="Slicer_1", cache="Slicer_1"),
            Slicer(name="Slicer_2", cache="Slicer_2"),
        ])
        assert len(slicers) == 2

    def test_from_xml(self, SlicerList):
        src = """
        <slicerList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <slicer name="Slicer_A" cache="Slicer_A" caption="Column A"/>
            <slicer name="Slicer_B" cache="Slicer_B" caption="Column B"/>
        </slicerList>
        """
        node = fromstring(src)
        slicer_list = SlicerList.from_tree(node)
        assert len(slicer_list) == 2
        assert slicer_list.slicer[0].name == "Slicer_A"
        assert slicer_list.slicer[1].name == "Slicer_B"

    def test_bool(self, SlicerList, Slicer):
        empty = SlicerList()
        assert not empty

        with_slicers = SlicerList(slicer=[
            Slicer(name="Test", cache="Test")
        ])
        assert with_slicers

    def test_path(self, SlicerList):
        sl = SlicerList()
        sl._id = 1
        assert sl.path == "/xl/slicers/slicer1.xml"


class TestSlicerCacheDefinition:

    def test_ctor(self, SlicerCacheDefinition):
        cache = SlicerCacheDefinition(
            name="Slicer_Category",
            sourceName="Category"
        )
        assert cache.name == "Slicer_Category"
        assert cache.sourceName == "Category"

    def test_path(self, SlicerCacheDefinition):
        cache = SlicerCacheDefinition(name="Test")
        cache._id = 1
        assert cache.path == "/xl/slicerCaches/slicerCache1.xml"

    def test_from_xml(self, SlicerCacheDefinition):
        src = """
        <slicerCacheDefinition
            name="Slicer_Region"
            sourceName="Region"/>
        """
        node = fromstring(src)
        cache = SlicerCacheDefinition.from_tree(node)
        assert cache.name == "Slicer_Region"
        assert cache.sourceName == "Region"
