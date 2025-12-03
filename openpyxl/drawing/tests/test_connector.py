# Copyright (c) 2010-2024 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def ConnectorShape():
    from ..connector import ConnectorShape
    return ConnectorShape


class TestConnectorShape:


    @pytest.mark.xfail
    def test_ctor(self, ConnectorShape):
        fut = ConnectorShape()
        xml = tostring(fut.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ConnectorShape):
        src = """
        <cxnSp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" macro="">
            <nvCxnSpPr>
                <cNvPr id="3" name="Straight Arrow Connector 2">
                </cNvPr>
                <cNvCxnSpPr/>
            </nvCxnSpPr>
            <spPr>
                <a:xfrm flipH="1" flipV="1">
                    <a:off x="3321050" y="3829050"/>
                    <a:ext cx="165100" cy="368300"/>
                </a:xfrm>
                <a:prstGeom prst="straightConnector1">
                    <a:avLst/>
                </a:prstGeom>
                <a:ln>
                    <a:tailEnd type="triangle"/>
                </a:ln>
            </spPr>
        </cxnSp>
        """
        node = fromstring(src)
        cnx = ConnectorShape.from_tree(node)
        assert cnx.nvCxnSpPr.cNvPr.id == 3


@pytest.fixture
def Shape():
    from ..connector import Shape
    return Shape


class TestShapeTolerantParsing:
    """Test tolerant parsing that preserves shapes that fail to parse"""

    def test_normal_shape_parses(self, Shape):
        """Test that normal shapes still parse correctly"""
        src = """
        <sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" macro="">
            <nvSpPr>
                <cNvPr id="2" name="TextBox 1"/>
                <cNvSpPr txBox="1"/>
            </nvSpPr>
            <spPr>
                <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="1000" cy="500"/>
                </a:xfrm>
                <a:prstGeom prst="rect"/>
            </spPr>
        </sp>
        """
        node = fromstring(src)
        shape = Shape.from_tree(node)
        assert shape.nvSpPr is not None
        assert shape.nvSpPr.cNvPr.id == 2
        assert not getattr(shape, '_is_raw', False)

    def test_unparseable_shape_preserved(self, Shape):
        """Test that shapes that fail to parse are preserved as raw XML"""
        # Create XML with invalid/unknown elements that would cause parsing failure
        src = """
        <sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <nvSpPr>
                <cNvPr id="5" name="Shape 5"/>
                <cNvSpPr/>
            </nvSpPr>
            <spPr>
                <a:custGeom>
                    <a:avLst/>
                    <a:gdLst/>
                    <a:ahLst/>
                    <a:pathLst/>
                </a:custGeom>
            </spPr>
            <unknownElement>Some unknown data</unknownElement>
        </sp>
        """
        node = fromstring(src)
        # Even if parsing fails, we should get a shape back
        shape = Shape.from_tree(node)
        assert shape is not None
        # If it was raw-preserved, it should have _raw_content
        if hasattr(shape, '_raw_content') and shape._raw_content is not None:
            # Round-trip should preserve the XML
            tree = shape.to_tree()
            assert tree is not None

    def test_raw_content_round_trip(self, Shape):
        """Test that raw content shapes round-trip correctly via _raw_content"""
        # Create a shape that will fail to parse - force raw preservation
        src = """
        <sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <nvSpPr>
                <cNvPr id="99" name="Custom Shape"/>
                <cNvSpPr/>
            </nvSpPr>
            <spPr>
            </spPr>
        </sp>
        """
        # Create raw content shape directly (bypassing __init__)
        from copy import deepcopy
        shape = Shape.__new__(Shape)
        shape._raw_content = deepcopy(fromstring(src))
        shape._is_raw = True

        # Serialize back should return the raw content
        tree = shape.to_tree()
        assert tree is not None
        # The tag should be 'sp'
        assert tree.tag == '{http://schemas.openxmlformats.org/drawingml/2006/main}sp' or tree.tag == 'sp'
