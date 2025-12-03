# Copyright (c) 2010-2024 openpyxl

import pytest

from .. import Strict


@pytest.fixture
def UniversalMeasure():
    from ..excel import UniversalMeasure

    class Dummy(Strict):

        value = UniversalMeasure()

    return Dummy()


class TestUniversalMeasure:

    @pytest.mark.parametrize("value",
                             ["24.73mm", "0cm", "24pt", '999pc', "50pi"]
                             )
    def test_valid(self, UniversalMeasure, value):
        UniversalMeasure.value = value
        assert UniversalMeasure.value == value

    @pytest.mark.parametrize("value",
                             [24.73, '24.73zz', "24.73 mm", None, "-24.73cm"]
                             )
    def test_invalid(self, UniversalMeasure, value):
        with pytest.raises(ValueError):
            UniversalMeasure.value = "{0}".format(value)


@pytest.fixture
def HexBinary():
    from ..excel import HexBinary

    class Dummy(Strict):

        value = HexBinary()

    return Dummy()


class TestHexBinary:

    @pytest.mark.parametrize("value",
                             ["aa35efd", "AABBCCDD"]
                             )
    def test_valid(self, HexBinary, value):
        HexBinary.value = value
        assert HexBinary.value == value


    @pytest.mark.parametrize("value",
                             ["GGII", "35.5"]
                             )
    def test_invalid(self, HexBinary, value):
        with pytest.raises(ValueError):
            HexBinary.value = value


@pytest.fixture
def TextPoint():
    from ..excel import TextPoint

    class Dummy(Strict):

        value = TextPoint()

    return Dummy()


class TestTextPoint:

    @pytest.mark.parametrize("value",
                             [-400000, "400000", 0]
                             )
    def test_valid(self, TextPoint, value):
        TextPoint.value = value
        assert TextPoint.value == int(value)

    def test_invalid_value(self, TextPoint):
        with pytest.raises(ValueError):
            TextPoint.value = -400001

    def test_invalid_type(self, TextPoint):
        with pytest.raises(TypeError):
            TextPoint.value = "40pt"


@pytest.fixture
def Percentage():
    from ..excel import Percentage

    class Dummy(Strict):

        value = Percentage()

    return Dummy()


class TestPercentage:

    @pytest.mark.parametrize("input, value",
                             [
                                 ("15%", 15000),
                                 (1500, 1500),
                                 ("15.5%", 15500),
                              ]
                             )
    def test_valid(self, Percentage, input, value):
        Percentage.value = value
        assert Percentage.value == value


    @pytest.mark.parametrize("value",
                             ["2000000", "-1000001",]
                             )
    def test_invalid(self, Percentage, value):
        with pytest.raises(ValueError):
            Percentage.value = value


@pytest.fixture
def Guid():
    from ..excel import Guid

    class Dummy(Strict):
        value = Guid()

    return Dummy()


class TestGuid():
    @pytest.mark.parametrize("value",
                             ["{00000000-5BD2-4BC8-9F70-7020E1357FB2}"]
                             )
    def test_valid(self, Guid, value):
        Guid.value = value
        assert Guid.value == value

    @pytest.mark.parametrize("value",
                             ["{00000000-5BD2-4BC8-9F70-7020E1357FB2"]
                             )
    def test_valid(self, Guid, value):
        with pytest.raises(ValueError):
            Guid.value = value


@pytest.fixture
def Base64Binary():
    from ..excel import Base64Binary

    class Dummy(Strict):
        value = Base64Binary()

    return Dummy()


class TestBase64Binary():
    @pytest.mark.parametrize("value",
                             ["9oN7nWkCAyEZib1RomSJTjmPpCY="]
                             )
    def test_valid(self, Base64Binary, value):
        Base64Binary.value = value
        assert Base64Binary.value == value

    @pytest.mark.parametrize("value",
                             ["==0F"]
                             )
    def test_valid(self, Base64Binary, value):
        with pytest.raises(ValueError):
            Base64Binary.value = value


@pytest.fixture
def CellRange():
    from ..excel import CellRange

    class Dummy(Strict):
        value = CellRange()

    return Dummy()


class TestCellRange():

    @pytest.mark.parametrize("value",
                             ["A1",
                              "A1:H5",
                              "A:B",
                              ]
                             )
    def test_valid(self, CellRange, value):
        CellRange.value = value
        assert CellRange.value == value


    @pytest.mark.parametrize("value",
                             ["A1:",
                              "A1:5",
                              "A1:B4:C7"
                              ]
                             )
    def test_invalid(self, CellRange, value):
        with pytest.raises(ValueError):
            CellRange.value = value


class TestExtension:
    """Tests for Extension class XML preservation"""

    def test_from_tree_preserves_content(self):
        from openpyxl.xml.functions import fromstring, tostring
        from ..excel import Extension

        # Simulate a sparkline extension with complex content
        src = """
        <ext uri="{05C60535-1F16-4FD2-B633-F4F36F0B64E0}"
             xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
             xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:sparklineGroups>
                <x14:sparklineGroup type="line">
                    <x14:sparklines>
                        <x14:sparkline>
                            <x14:f>Sheet1!A1:A10</x14:f>
                            <x14:sqref>B1</x14:sqref>
                        </x14:sparkline>
                    </x14:sparklines>
                </x14:sparklineGroup>
            </x14:sparklineGroups>
        </ext>
        """
        node = fromstring(src)
        ext = Extension.from_tree(node)

        # Check URI is extracted
        assert ext.uri == "{05C60535-1F16-4FD2-B633-F4F36F0B64E0}"

        # Check content is preserved
        assert ext._content is not None

        # Round-trip: to_tree should return equivalent content
        result = ext.to_tree()
        assert result.get('uri') == ext.uri

        # The inner content should be preserved
        result_str = tostring(result).decode()
        assert 'sparklineGroups' in result_str
        assert 'Sheet1!A1:A10' in result_str

    def test_to_tree_without_content(self):
        from openpyxl.xml.functions import tostring
        from ..excel import Extension

        # Programmatically created extension (no preserved content)
        ext = Extension(uri="{TEST-GUID}")
        result = ext.to_tree()

        assert result.get('uri') == "{TEST-GUID}"

    def test_content_is_deep_copy(self):
        from openpyxl.xml.functions import fromstring
        from ..excel import Extension

        src = '<ext uri="{TEST}"><child attr="value"/></ext>'
        node = fromstring(src)
        ext = Extension.from_tree(node)

        # Modify original node
        node.set('uri', 'MODIFIED')

        # Extension should not be affected
        assert ext.uri == "{TEST}"
        assert ext._content.get('uri') == "{TEST}"


class TestExtensionList:
    """Tests for ExtensionList round-trip preservation"""

    def test_from_tree_preserves_all_extensions(self):
        from openpyxl.xml.functions import fromstring, tostring
        from ..excel import ExtensionList

        src = """
        <extLst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <ext uri="{05C60535-1F16-4FD2-B633-F4F36F0B64E0}">
                <sparklineData>test1</sparklineData>
            </ext>
            <ext uri="{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}">
                <slicerData>test2</slicerData>
            </ext>
        </extLst>
        """
        node = fromstring(src)
        extLst = ExtensionList.from_tree(node)

        # Should have 2 extensions
        assert len(extLst.ext) == 2

        # Each should have preserved content
        assert extLst.ext[0].uri == "{05C60535-1F16-4FD2-B633-F4F36F0B64E0}"
        assert extLst.ext[0]._content is not None
        assert extLst.ext[1].uri == "{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}"
        assert extLst.ext[1]._content is not None

        # Round-trip
        result = extLst.to_tree()
        result_str = tostring(result).decode()
        assert 'sparklineData' in result_str
        assert 'slicerData' in result_str
        assert 'test1' in result_str
        assert 'test2' in result_str
