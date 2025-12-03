# Copyright (c) 2010-2024 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def Sparkline():
    from ..sparkline import Sparkline
    return Sparkline


@pytest.fixture
def SparklineGroup():
    from ..sparkline import SparklineGroup
    return SparklineGroup


@pytest.fixture
def SparklineGroups():
    from ..sparkline import SparklineGroups
    return SparklineGroups


@pytest.fixture
def SparklineColor():
    from ..sparkline import SparklineColor
    return SparklineColor


class TestSparkline:

    def test_ctor(self, Sparkline):
        spark = Sparkline(f="Sheet1!A1:A10", sqref="B1")
        assert spark.f == "Sheet1!A1:A10"
        assert spark.sqref == "B1"

    def test_to_tree(self, Sparkline):
        spark = Sparkline(f="Sheet1!A1:A5", sqref="C1")
        xml = tostring(spark.to_tree())
        expected = """
        <sparkline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <f>Sheet1!A1:A5</f>
            <sqref>C1</sqref>
        </sparkline>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


class TestSparklineColor:

    def test_ctor(self, SparklineColor):
        color = SparklineColor(theme=4, tint=0.5)
        assert color.theme == 4
        assert color.tint == 0.5

    def test_rgb(self, SparklineColor):
        color = SparklineColor(rgb="FF0000")
        assert color.rgb == "FF0000"


class TestSparklineGroup:

    def test_ctor(self, SparklineGroup, Sparkline):
        group = SparklineGroup(
            type="line",
            displayEmptyCellsAs="gap",
            markers=True,
            sparklines=[Sparkline(f="Sheet1!A1:A5", sqref="B1")]
        )
        assert group.type == "line"
        assert group.displayEmptyCellsAs == "gap"
        assert group.markers is True
        assert len(group.sparklines) == 1

    def test_column_type(self, SparklineGroup):
        group = SparklineGroup(type="column")
        assert group.type == "column"

    def test_stacked_type(self, SparklineGroup):
        """Win/loss sparklines use 'stacked' type"""
        group = SparklineGroup(type="stacked")
        assert group.type == "stacked"


class TestSparklineGroups:

    def test_ctor(self, SparklineGroups, SparklineGroup, Sparkline):
        groups = SparklineGroups(
            sparklineGroup=[
                SparklineGroup(
                    type="line",
                    sparklines=[Sparkline(f="A1:A10", sqref="B1")]
                )
            ]
        )
        assert len(groups) == 1

    def test_access(self, SparklineGroups, SparklineGroup, Sparkline):
        groups = SparklineGroups(
            sparklineGroup=[
                SparklineGroup(type="line"),
                SparklineGroup(type="column"),
            ]
        )
        types = [g.type for g in groups.sparklineGroup]
        assert types == ["line", "column"]

    def test_append(self, SparklineGroups, SparklineGroup):
        groups = SparklineGroups()
        groups.append(SparklineGroup(type="line"))
        assert len(groups) == 1

    def test_from_xml(self, SparklineGroups):
        """Test parsing SparklineGroups from XML"""
        src = """
        <sparklineGroups xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <sparklineGroup type="line" displayEmptyCellsAs="gap">
                <sparklines>
                    <sparkline>
                        <f>Sheet1!A1:A10</f>
                        <sqref>B1</sqref>
                    </sparkline>
                    <sparkline>
                        <f>Sheet1!A1:A10</f>
                        <sqref>B2</sqref>
                    </sparkline>
                </sparklines>
            </sparklineGroup>
        </sparklineGroups>
        """
        node = fromstring(src)
        groups = SparklineGroups.from_tree(node)
        assert len(groups) == 1
        group = groups.sparklineGroup[0]
        assert group.type == "line"
        assert group.displayEmptyCellsAs == "gap"
        assert len(group.sparklines) == 2
        spark = group.sparklines[0]
        assert spark.f == "Sheet1!A1:A10"
        assert spark.sqref == "B1"

    def test_round_trip(self, SparklineGroups, SparklineGroup, Sparkline):
        """Test that sparklines can be serialized and deserialized"""
        original = SparklineGroups(
            sparklineGroup=[
                SparklineGroup(
                    type="column",
                    displayEmptyCellsAs="zero",
                    sparklines=[
                        Sparkline(f="Sheet1!C1:C5", sqref="D1"),
                        Sparkline(f="Sheet1!C1:C5", sqref="D2"),
                    ]
                )
            ]
        )

        # Serialize
        tree = original.to_tree()
        xml = tostring(tree)

        # Deserialize
        parsed = SparklineGroups.from_tree(fromstring(xml))

        # Verify
        assert len(parsed) == 1
        group = parsed.sparklineGroup[0]
        assert group.type == "column"
        assert len(group.sparklines) == 2


class TestSparklineReader:
    """Test sparkline parsing from worksheet extensions"""

    def test_parse_sparklines_from_extension(self):
        """Test that sparklines can be parsed from extLst"""
        from openpyxl.worksheet._reader import WorkSheetParser
        from io import BytesIO
        import warnings

        src = b"""
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                   xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <sheetData/>
            <extLst>
                <ext uri="{05C60535-1F16-4FD2-B633-F4F36F0B64E0}">
                    <x14:sparklineGroups>
                        <x14:sparklineGroup type="line" displayEmptyCellsAs="gap">
                            <x14:sparklines>
                                <x14:sparkline>
                                    <x14:f>Sheet1!A1:A10</x14:f>
                                    <x14:sqref>B1</x14:sqref>
                                </x14:sparkline>
                            </x14:sparklines>
                        </x14:sparklineGroup>
                    </x14:sparklineGroups>
                </ext>
            </extLst>
        </worksheet>
        """
        parser = WorkSheetParser(BytesIO(src), [], False, None, [], [], False)

        # Parse the worksheet
        for _ in parser.parse():
            pass

        # Check sparklines were parsed
        assert parser.sparklines is not None
        assert len(parser.sparklines) == 1
        group = parser.sparklines.sparklineGroup[0]
        assert group.type == "line"
        assert len(group.sparklines) == 1
        assert group.sparklines[0].f == "Sheet1!A1:A10"
        assert group.sparklines[0].sqref == "B1"
