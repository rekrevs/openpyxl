# Copyright (c) 2010-2024 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring

from ..chartex import (
    ChartExSpace,
    CHARTEX_NS,
    CHARTEX_TYPE,
    is_chartex,
)


class TestChartExSpace:

    def test_ctor(self):
        chartex = ChartExSpace()
        assert chartex._content is None
        assert not chartex

    def test_from_tree(self):
        src = """
        <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
            <cx:chart>
                <cx:title>
                    <cx:tx>Test Chart</cx:tx>
                </cx:title>
                <cx:plotArea>
                    <cx:plotAreaRegion>
                        <cx:series>
                            <cx:dataId val="0"/>
                        </cx:series>
                    </cx:plotAreaRegion>
                </cx:plotArea>
            </cx:chart>
        </cx:chartSpace>
        """
        node = fromstring(src)
        chartex = ChartExSpace.from_tree(node)
        assert chartex._content is not None
        assert chartex

    def test_to_tree_preserves_content(self):
        src = """
        <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
            <cx:chart>
                <cx:title>
                    <cx:tx>Waterfall Chart</cx:tx>
                </cx:title>
            </cx:chart>
        </cx:chartSpace>
        """
        node = fromstring(src)
        chartex = ChartExSpace.from_tree(node)

        # Serialize and deserialize
        tree = chartex.to_tree()
        assert tree is not None
        assert "chartSpace" in tree.tag

    def test_path(self):
        chartex = ChartExSpace()
        chartex._id = 1
        assert chartex.path == "/xl/charts/chartEx1.xml"

    def test_mime_type(self):
        assert ChartExSpace.mime_type == CHARTEX_TYPE

    def test_empty_chartex(self):
        chartex = ChartExSpace()
        assert not chartex

        # to_tree should return an empty element
        tree = chartex.to_tree()
        assert tree is not None


class TestIsChartEx:

    def test_is_chartex_true(self):
        src = """
        <cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
        </cx:chartSpace>
        """
        node = fromstring(src)
        assert is_chartex(node)

    def test_is_chartex_false(self):
        src = """
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        </c:chartSpace>
        """
        node = fromstring(src)
        assert not is_chartex(node)

    def test_is_chartex_none(self):
        assert not is_chartex(None)
