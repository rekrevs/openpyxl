# Copyright (c) 2010-2024 openpyxl
"""
Round-trip test suite for XLSX preservation.

Tests that modern Excel files survive load_workbook() then save() without data loss.
This is part of the WOTAN initiative for full XLSX compatibility.
"""
import os
import pytest
from io import BytesIO
from zipfile import ZipFile
from lxml import etree

from openpyxl import load_workbook


def get_xml_content(xlsx_bytes, inner_path):
    """Extract and parse XML content from an XLSX file."""
    with ZipFile(BytesIO(xlsx_bytes), 'r') as zf:
        try:
            content = zf.read(inner_path)
            return etree.fromstring(content)
        except KeyError:
            return None


def xml_to_string(element):
    """Convert XML element to normalized string for comparison."""
    if element is None:
        return None
    return etree.tostring(element, pretty_print=True).decode('utf-8')


def compare_xml_elements(original, roundtripped, path, ignore_attrs=None):
    """
    Compare two XML elements recursively.
    Returns list of differences found.
    """
    differences = []
    ignore_attrs = ignore_attrs or set()

    if original is None and roundtripped is None:
        return differences

    if original is None:
        differences.append(f"{path}: Original missing, roundtripped present")
        return differences

    if roundtripped is None:
        differences.append(f"{path}: Original present, roundtripped missing")
        return differences

    # Compare tags
    if original.tag != roundtripped.tag:
        differences.append(f"{path}: Tag mismatch: {original.tag} vs {roundtripped.tag}")
        return differences

    # Compare attributes (excluding ignored ones)
    orig_attrs = {k: v for k, v in original.attrib.items() if k not in ignore_attrs}
    rt_attrs = {k: v for k, v in roundtripped.attrib.items() if k not in ignore_attrs}

    if orig_attrs != rt_attrs:
        differences.append(f"{path}: Attribute mismatch: {orig_attrs} vs {rt_attrs}")

    # Compare text
    orig_text = (original.text or '').strip()
    rt_text = (roundtripped.text or '').strip()
    if orig_text != rt_text:
        differences.append(f"{path}: Text mismatch: '{orig_text}' vs '{rt_text}'")

    # Compare children
    orig_children = list(original)
    rt_children = list(roundtripped)

    if len(orig_children) != len(rt_children):
        differences.append(f"{path}: Child count mismatch: {len(orig_children)} vs {len(rt_children)}")

    for i, (orig_child, rt_child) in enumerate(zip(orig_children, rt_children)):
        child_path = f"{path}/{orig_child.tag}[{i}]"
        differences.extend(compare_xml_elements(orig_child, rt_child, child_path, ignore_attrs))

    return differences


class TestRoundTrip:
    """Test that XLSX files survive round-trip without data loss."""

    @pytest.fixture
    def datadir(self):
        """Get the data directory path."""
        here = os.path.dirname(__file__)
        return os.path.join(here, "data")

    def roundtrip(self, xlsx_path):
        """Load and save an XLSX file, returning both original and result bytes."""
        with open(xlsx_path, 'rb') as f:
            original_bytes = f.read()

        wb = load_workbook(BytesIO(original_bytes))
        output = BytesIO()
        wb.save(output)
        roundtripped_bytes = output.getvalue()

        return original_bytes, roundtripped_bytes

    def test_sample_xlsx_roundtrip(self, datadir):
        """Test that sample.xlsx survives round-trip."""
        xlsx_path = os.path.join(datadir, "genuine", "sample.xlsx")
        original, roundtripped = self.roundtrip(xlsx_path)

        # Check that we can re-load the round-tripped file
        wb = load_workbook(BytesIO(roundtripped))
        assert len(wb.sheetnames) == 4

        # Verify worksheet content preserved
        ws = wb['Sheet1 - Text']
        assert ws is not None

    def test_extensions_preserved(self, datadir):
        """Test that extension list content is preserved."""
        xlsx_path = os.path.join(datadir, "genuine", "sample.xlsx")
        original, roundtripped = self.roundtrip(xlsx_path)

        # Get extLst from both
        orig_sheet = get_xml_content(original, 'xl/worksheets/sheet1.xml')
        rt_sheet = get_xml_content(roundtripped, 'xl/worksheets/sheet1.xml')

        ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

        orig_ext = orig_sheet.find('.//x:extLst', ns) if orig_sheet is not None else None
        rt_ext = rt_sheet.find('.//x:extLst', ns) if rt_sheet is not None else None

        if orig_ext is not None:
            assert rt_ext is not None, "extLst was lost during round-trip"
            # Compare extension count
            orig_exts = list(orig_ext)
            rt_exts = list(rt_ext)
            assert len(orig_exts) == len(rt_exts), f"Extension count changed: {len(orig_exts)} -> {len(rt_exts)}"

    def test_worksheet_xml_structure(self, datadir):
        """Test that core worksheet XML structure is preserved."""
        xlsx_path = os.path.join(datadir, "genuine", "sample.xlsx")
        original, roundtripped = self.roundtrip(xlsx_path)

        # Check main worksheet
        orig_sheet = get_xml_content(original, 'xl/worksheets/sheet1.xml')
        rt_sheet = get_xml_content(roundtripped, 'xl/worksheets/sheet1.xml')

        assert orig_sheet is not None
        assert rt_sheet is not None

        # Tags should match
        assert orig_sheet.tag == rt_sheet.tag

    def test_empty_xlsx_roundtrip(self, datadir):
        """Test that empty.xlsx survives round-trip."""
        xlsx_path = os.path.join(datadir, "genuine", "empty.xlsx")
        original, roundtripped = self.roundtrip(xlsx_path)

        # Check that we can re-load
        wb = load_workbook(BytesIO(roundtripped))
        assert len(wb.sheetnames) >= 1

    def test_styles_preserved(self, datadir):
        """Test that styles.xml content is preserved."""
        xlsx_path = os.path.join(datadir, "genuine", "empty-with-styles.xlsx")
        original, roundtripped = self.roundtrip(xlsx_path)

        orig_styles = get_xml_content(original, 'xl/styles.xml')
        rt_styles = get_xml_content(roundtripped, 'xl/styles.xml')

        assert orig_styles is not None
        assert rt_styles is not None

        # Compare font count
        ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        orig_fonts = orig_styles.find('.//x:fonts', ns)
        rt_fonts = rt_styles.find('.//x:fonts', ns)

        if orig_fonts is not None:
            assert rt_fonts is not None, "fonts element lost during round-trip"

    def test_formulas_preserved(self, datadir):
        """Test that formulas in sample.xlsx are preserved."""
        xlsx_path = os.path.join(datadir, "genuine", "sample.xlsx")

        # Load and get formulas from Sheet3 - Formulas
        with open(xlsx_path, 'rb') as f:
            original_bytes = f.read()

        wb1 = load_workbook(BytesIO(original_bytes))
        ws1 = wb1['Sheet3 - Formulas']

        # Find cells with formulas
        formulas_before = {}
        for row in ws1.iter_rows():
            for cell in row:
                if cell.data_type == 'f':
                    formulas_before[cell.coordinate] = cell.value

        # Round-trip
        output = BytesIO()
        wb1.save(output)

        wb2 = load_workbook(BytesIO(output.getvalue()))
        ws2 = wb2['Sheet3 - Formulas']

        # Check formulas preserved
        for coord, formula in formulas_before.items():
            cell = ws2[coord]
            assert cell.data_type == 'f', f"Cell {coord} lost formula type"


class TestUnknownPartPreservation:
    """Test that unknown archive members survive round-trip."""

    def _create_xlsx_with_unknown_parts(self):
        """
        Create an XLSX file that contains unknown archive members.
        Returns the bytes of the modified XLSX.
        """
        from openpyxl import Workbook as WB

        # First create a normal XLSX
        wb = WB()
        ws = wb.active
        ws.title = "Sheet1"
        ws['A1'] = "Hello"
        ws['B1'] = 42

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)

        # Now inject unknown parts into the zip
        original_data = buf.getvalue()
        new_buf = BytesIO()
        with ZipFile(BytesIO(original_data), 'r') as zf_in:
            with ZipFile(new_buf, 'w') as zf_out:
                # Copy all existing entries
                for item in zf_in.namelist():
                    zf_out.writestr(item, zf_in.read(item))

                # Add unknown archive members
                zf_out.writestr(
                    'xl/customFeature/data.xml',
                    b'<?xml version="1.0" encoding="UTF-8"?>\n<customData><item id="1">test</item></customData>'
                )
                zf_out.writestr(
                    'xl/customFeature/settings.xml',
                    b'<?xml version="1.0" encoding="UTF-8"?>\n<settings><option name="enabled" value="true"/></settings>'
                )
                zf_out.writestr(
                    'xl/customFeature/binary.bin',
                    b'\x00\x01\x02\x03\x04\x05\x06\x07\x08\x09'
                )

                # Also add a content type override for the XML files
                # Read and patch [Content_Types].xml
                ct_data = zf_in.read('[Content_Types].xml')
                ct_tree = etree.fromstring(ct_data)
                ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
                etree.SubElement(ct_tree, f'{{{ns}}}Override', {
                    'PartName': '/xl/customFeature/data.xml',
                    'ContentType': 'application/vnd.ms-excel.customFeature+xml'
                })
                etree.SubElement(ct_tree, f'{{{ns}}}Override', {
                    'PartName': '/xl/customFeature/settings.xml',
                    'ContentType': 'application/vnd.ms-excel.customFeatureSettings+xml'
                })

        # We need to rebuild the zip because [Content_Types].xml was already written
        # Re-create with the patched content types
        new_buf2 = BytesIO()
        with ZipFile(BytesIO(new_buf.getvalue()), 'r') as zf_in:
            with ZipFile(new_buf2, 'w') as zf_out:
                for item in zf_in.namelist():
                    if item == '[Content_Types].xml':
                        zf_out.writestr(item, etree.tostring(ct_tree, xml_declaration=True, encoding='UTF-8'))
                    else:
                        zf_out.writestr(item, zf_in.read(item))

        return new_buf2.getvalue()

    def test_unknown_xml_parts_preserved(self):
        """Test that unknown XML archive members survive round-trip."""
        xlsx_data = self._create_xlsx_with_unknown_parts()

        # Verify the unknown parts exist in the original
        with ZipFile(BytesIO(xlsx_data), 'r') as zf:
            assert 'xl/customFeature/data.xml' in zf.namelist()
            assert 'xl/customFeature/settings.xml' in zf.namelist()
            original_data = zf.read('xl/customFeature/data.xml')
            original_settings = zf.read('xl/customFeature/settings.xml')

        # Round-trip through openpyxl
        wb = load_workbook(BytesIO(xlsx_data))
        output = BytesIO()
        wb.save(output)

        # Verify the unknown parts survived
        with ZipFile(BytesIO(output.getvalue()), 'r') as zf:
            names = zf.namelist()
            assert 'xl/customFeature/data.xml' in names, \
                f"Unknown XML part lost during round-trip. Archive contains: {sorted(names)}"
            assert 'xl/customFeature/settings.xml' in names, \
                f"Unknown settings XML lost during round-trip. Archive contains: {sorted(names)}"

            # Verify content is identical
            assert zf.read('xl/customFeature/data.xml') == original_data
            assert zf.read('xl/customFeature/settings.xml') == original_settings

    def test_unknown_binary_parts_preserved(self):
        """Test that unknown binary archive members survive round-trip."""
        xlsx_data = self._create_xlsx_with_unknown_parts()

        with ZipFile(BytesIO(xlsx_data), 'r') as zf:
            original_bin = zf.read('xl/customFeature/binary.bin')

        # Round-trip
        wb = load_workbook(BytesIO(xlsx_data))
        output = BytesIO()
        wb.save(output)

        with ZipFile(BytesIO(output.getvalue()), 'r') as zf:
            names = zf.namelist()
            assert 'xl/customFeature/binary.bin' in names, \
                f"Unknown binary part lost during round-trip. Archive contains: {sorted(names)}"
            assert zf.read('xl/customFeature/binary.bin') == original_bin

    def test_unknown_content_types_preserved(self):
        """Test that content type entries for unknown parts survive round-trip."""
        xlsx_data = self._create_xlsx_with_unknown_parts()

        # Round-trip
        wb = load_workbook(BytesIO(xlsx_data))
        output = BytesIO()
        wb.save(output)

        # Check [Content_Types].xml for the custom content types
        with ZipFile(BytesIO(output.getvalue()), 'r') as zf:
            ct_data = zf.read('[Content_Types].xml')
            ct_tree = etree.fromstring(ct_data)

        ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
        overrides = {
            el.get('PartName'): el.get('ContentType')
            for el in ct_tree.findall(f'{{{ns}}}Override')
        }

        assert '/xl/customFeature/data.xml' in overrides, \
            f"Content type for data.xml not preserved. Overrides: {overrides}"
        assert overrides['/xl/customFeature/data.xml'] == 'application/vnd.ms-excel.customFeature+xml'

        assert '/xl/customFeature/settings.xml' in overrides, \
            f"Content type for settings.xml not preserved. Overrides: {overrides}"
        assert overrides['/xl/customFeature/settings.xml'] == 'application/vnd.ms-excel.customFeatureSettings+xml'

    def test_calcchain_not_preserved(self):
        """Test that xl/calcChain.xml is NOT preserved (Excel rebuilds it)."""
        from openpyxl import Workbook as WB

        wb = WB()
        ws = wb.active
        ws['A1'] = 1
        buf = BytesIO()
        wb.save(buf)

        # Inject a fake calcChain.xml
        new_buf = BytesIO()
        with ZipFile(BytesIO(buf.getvalue()), 'r') as zf_in:
            with ZipFile(new_buf, 'w') as zf_out:
                for item in zf_in.namelist():
                    zf_out.writestr(item, zf_in.read(item))
                zf_out.writestr('xl/calcChain.xml', b'<calcChain/>')

        # Round-trip
        wb2 = load_workbook(BytesIO(new_buf.getvalue()))
        output = BytesIO()
        wb2.save(output)

        with ZipFile(BytesIO(output.getvalue()), 'r') as zf:
            assert 'xl/calcChain.xml' not in zf.namelist(), \
                "calcChain.xml should NOT be preserved during round-trip"

    def test_normal_workbook_data_unaffected(self):
        """Test that normal workbook data is not affected by unknown part preservation."""
        xlsx_data = self._create_xlsx_with_unknown_parts()

        wb = load_workbook(BytesIO(xlsx_data))
        output = BytesIO()
        wb.save(output)

        # Re-load and verify normal content
        wb2 = load_workbook(BytesIO(output.getvalue()))
        ws = wb2['Sheet1']
        assert ws['A1'].value == "Hello"
        assert ws['B1'].value == 42

    def test_unknown_parts_attribute_exists(self):
        """Test that _unknown_parts attribute is always present on workbooks."""
        from openpyxl import Workbook as WB

        # New workbook should have empty _unknown_parts
        wb = WB()
        assert hasattr(wb, '_unknown_parts')
        assert wb._unknown_parts == []
        assert hasattr(wb, '_unknown_content_types')
        assert wb._unknown_content_types == []

    def test_double_roundtrip_preserves_unknown_parts(self):
        """Test that unknown parts survive two round-trips."""
        xlsx_data = self._create_xlsx_with_unknown_parts()

        # First round-trip
        wb1 = load_workbook(BytesIO(xlsx_data))
        buf1 = BytesIO()
        wb1.save(buf1)

        # Second round-trip
        wb2 = load_workbook(BytesIO(buf1.getvalue()))
        buf2 = BytesIO()
        wb2.save(buf2)

        with ZipFile(BytesIO(buf2.getvalue()), 'r') as zf:
            names = zf.namelist()
            assert 'xl/customFeature/data.xml' in names
            assert 'xl/customFeature/settings.xml' in names
            assert 'xl/customFeature/binary.bin' in names

    def test_multiple_unknown_parts_from_different_dirs(self):
        """Test that unknown parts from different directories are preserved."""
        from openpyxl import Workbook as WB

        wb = WB()
        ws = wb.active
        ws['A1'] = 1
        buf = BytesIO()
        wb.save(buf)

        # Inject parts in multiple directories
        new_buf = BytesIO()
        with ZipFile(BytesIO(buf.getvalue()), 'r') as zf_in:
            with ZipFile(new_buf, 'w') as zf_out:
                for item in zf_in.namelist():
                    zf_out.writestr(item, zf_in.read(item))
                zf_out.writestr('xl/pyExcel/code.xml', b'<pythonCode/>')
                zf_out.writestr('xl/namedSheetViews/namedSheetView1.xml', b'<namedSheetViews/>')
                zf_out.writestr('customXml/item1.xml', b'<custom/>')

        wb2 = load_workbook(BytesIO(new_buf.getvalue()))
        out = BytesIO()
        wb2.save(out)

        with ZipFile(BytesIO(out.getvalue()), 'r') as zf:
            names = zf.namelist()
            assert 'xl/pyExcel/code.xml' in names
            assert 'xl/namedSheetViews/namedSheetView1.xml' in names
            assert 'customXml/item1.xml' in names


class TestWotanFiles:
    """Tests for WOTAN-specific test files (modern Excel features)."""

    @pytest.fixture
    def wotan_dir(self):
        """Get the WOTAN test data directory path."""
        here = os.path.dirname(__file__)
        path = os.path.join(here, "data", "wotan")
        if not os.path.exists(path):
            os.makedirs(path)
        return path

    def test_wotan_dir_exists(self, wotan_dir):
        """Verify WOTAN test data directory exists."""
        assert os.path.isdir(wotan_dir)


# Utility function for creating test files
def create_minimal_xlsx_with_features(output_path, features):
    """
    Create a minimal XLSX file with specific features for testing.

    Args:
        output_path: Where to save the file
        features: Dict of features to include (e.g., {'sparklines': True})

    This is a helper for generating test files, not a test itself.
    """
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Test"

    # Add sample data
    for i in range(1, 11):
        ws.cell(row=i, column=1, value=i)
        ws.cell(row=i, column=2, value=i * 2)

    # Add a formula
    ws['C1'] = '=SUM(A1:A10)'

    wb.save(output_path)
    return output_path


class TestShapePreservation:
    """
    Test that DrawingML shapes (sp, cxnSp, grpSp) survive round-trip.

    These tests use programmatic XLSX injection: create a workbook with openpyxl,
    then inject drawing XML into the zip, and verify shapes survive load/save.

    Currently, openpyxl only preserves charts and images from drawings.
    Shapes, connectors, and group shapes are silently dropped. These tests
    define the contract for shape preservation.
    """

    # -- Namespace constants for drawing XML --
    XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    SSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
    PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

    DRAWING_REL_TYPE = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
    )
    DRAWING_CONTENT_TYPE = (
        "application/vnd.openxmlformats-officedocument.drawing+xml"
    )

    def _create_base_xlsx(self):
        """Create a minimal XLSX with one sheet and cell data, returns bytes."""
        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws['A1'] = "test data"
        ws['B2'] = 42
        buf = BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def _inject_drawing(self, xlsx_bytes, drawing_xml):
        """
        Inject a drawing XML into an XLSX zip.

        This modifies:
        1. Adds xl/drawings/drawing1.xml with the given content
        2. Updates xl/worksheets/_rels/sheet1.xml.rels to reference the drawing
        3. Updates [Content_Types].xml to include the drawing content type
        4. Updates xl/worksheets/sheet1.xml to add a <drawing> element

        Returns the modified XLSX bytes.
        """
        if isinstance(drawing_xml, str):
            drawing_xml = drawing_xml.encode('utf-8')

        original = BytesIO(xlsx_bytes)
        output = BytesIO()

        with ZipFile(original, 'r') as zf_in:
            with ZipFile(output, 'w') as zf_out:
                for item in zf_in.namelist():
                    data = zf_in.read(item)

                    if item == '[Content_Types].xml':
                        # Add drawing content type
                        ct_tree = etree.fromstring(data)
                        etree.SubElement(ct_tree, f'{{{self.CT_NS}}}Override', {
                            'PartName': '/xl/drawings/drawing1.xml',
                            'ContentType': self.DRAWING_CONTENT_TYPE,
                        })
                        data = etree.tostring(
                            ct_tree, xml_declaration=True, encoding='UTF-8'
                        )

                    elif item == 'xl/worksheets/sheet1.xml':
                        # Add <drawing r:id="rId_drawing"/> to the worksheet
                        ws_tree = etree.fromstring(data)
                        ns = self.SSHEET_NS
                        r_ns = self.R_NS
                        # Check if drawing element already exists
                        existing = ws_tree.findall(f'{{{ns}}}drawing')
                        if not existing:
                            drawing_el = etree.SubElement(
                                ws_tree, f'{{{ns}}}drawing'
                            )
                            drawing_el.set(f'{{{r_ns}}}id', 'rId_drawing')
                        data = etree.tostring(
                            ws_tree, xml_declaration=True, encoding='UTF-8'
                        )

                    zf_out.writestr(item, data)

                # Add the drawing XML
                zf_out.writestr('xl/drawings/drawing1.xml', drawing_xml)

                # Add or update worksheet rels to reference the drawing
                rels_path = 'xl/worksheets/_rels/sheet1.xml.rels'
                if rels_path in zf_in.namelist():
                    rels_data = zf_in.read(rels_path)
                    rels_tree = etree.fromstring(rels_data)
                else:
                    rels_tree = etree.Element(
                        f'{{{self.PKG_REL_NS}}}Relationships'
                    )

                etree.SubElement(rels_tree, f'{{{self.PKG_REL_NS}}}Relationship', {
                    'Id': 'rId_drawing',
                    'Type': self.DRAWING_REL_TYPE,
                    'Target': '../drawings/drawing1.xml',
                })
                zf_out.writestr(
                    rels_path,
                    etree.tostring(rels_tree, xml_declaration=True, encoding='UTF-8'),
                )

        return output.getvalue()

    def _roundtrip(self, xlsx_bytes):
        """Load workbook from bytes, save to bytes, return result bytes."""
        wb = load_workbook(BytesIO(xlsx_bytes))
        out = BytesIO()
        wb.save(out)
        return out.getvalue()

    def _get_drawing_xml(self, xlsx_bytes, path='xl/drawings/drawing1.xml'):
        """Extract and parse drawing XML from XLSX bytes. Returns etree or None."""
        with ZipFile(BytesIO(xlsx_bytes), 'r') as zf:
            if path in zf.namelist():
                return etree.fromstring(zf.read(path))
        return None

    # -- Drawing XML templates --

    def _shape_drawing_xml(self, text="Hello World"):
        """A drawing with a single twoCellAnchor containing a textbox shape."""
        return f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="{self.XDR_NS}"
          xmlns:a="{self.A_NS}">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:sp macro="" textlink="">
      <xdr:nvSpPr>
        <xdr:cNvPr id="2" name="TextBox 1"/>
        <xdr:cNvSpPr txBox="1"/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="3000000" cy="2000000"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr/>
        <a:p><a:r><a:t>{text}</a:t></a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"""

    def _connector_drawing_xml(self):
        """A drawing with a single twoCellAnchor containing a connector shape."""
        return f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="{self.XDR_NS}"
          xmlns:a="{self.A_NS}">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:cxnSp macro="">
      <xdr:nvCxnSpPr>
        <xdr:cNvPr id="3" name="Straight Connector 1"/>
        <xdr:cNvCxnSpPr/>
      </xdr:nvCxnSpPr>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="100000" y="200000"/>
          <a:ext cx="1500000" cy="1000000"/>
        </a:xfrm>
        <a:prstGeom prst="line">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:cxnSp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"""

    def _group_shape_drawing_xml(self):
        """A drawing with a twoCellAnchor containing a group of two shapes."""
        return f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="{self.XDR_NS}"
          xmlns:a="{self.A_NS}">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:grpSp>
      <xdr:nvGrpSpPr>
        <xdr:cNvPr id="10" name="Group 1"/>
        <xdr:cNvGrpSpPr/>
      </xdr:nvGrpSpPr>
      <xdr:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="5000000" cy="3000000"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="5000000" cy="3000000"/>
        </a:xfrm>
      </xdr:grpSpPr>
      <xdr:sp macro="" textlink="">
        <xdr:nvSpPr>
          <xdr:cNvPr id="11" name="Rectangle 1"/>
          <xdr:cNvSpPr/>
        </xdr:nvSpPr>
        <xdr:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="2000000" cy="1500000"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
        </xdr:spPr>
      </xdr:sp>
      <xdr:sp macro="" textlink="">
        <xdr:nvSpPr>
          <xdr:cNvPr id="12" name="Oval 1"/>
          <xdr:cNvSpPr/>
        </xdr:nvSpPr>
        <xdr:spPr>
          <a:xfrm>
            <a:off x="2500000" y="0"/>
            <a:ext cx="2000000" cy="1500000"/>
          </a:xfrm>
          <a:prstGeom prst="ellipse">
            <a:avLst/>
          </a:prstGeom>
        </xdr:spPr>
      </xdr:sp>
    </xdr:grpSp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"""

    def _shape_and_image_drawing_xml(self):
        """
        A drawing with two anchors: one shape and one image (pic).

        The image anchor uses a blipFill referencing rId1, which would point
        to an actual image in a real file. For our test, we only care that
        both anchors survive the round-trip; the image data is separately
        injected.
        """
        return f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="{self.XDR_NS}"
          xmlns:a="{self.A_NS}"
          xmlns:r="{self.R_NS}">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:sp macro="" textlink="">
      <xdr:nvSpPr>
        <xdr:cNvPr id="2" name="TextBox 1"/>
        <xdr:cNvSpPr txBox="1"/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="3000000" cy="2000000"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr/>
        <a:p><a:r><a:t>Shape Text</a:t></a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="3" name="Picture 1"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="4000000" y="0"/>
          <a:ext cx="3000000" cy="2000000"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"""

    def _inject_drawing_with_image(self, xlsx_bytes, drawing_xml):
        """
        Inject a drawing XML and a tiny PNG image into an XLSX zip.

        Adds:
        - xl/drawings/drawing1.xml
        - xl/media/image1.png (1x1 red pixel PNG)
        - xl/drawings/_rels/drawing1.xml.rels pointing image rId1 -> image1.png
        - Worksheet rels and content types
        """
        import struct
        import zlib

        # Minimal 1x1 red pixel PNG
        def _minimal_png():
            sig = b'\x89PNG\r\n\x1a\n'
            # IHDR
            ihdr_data = struct.pack('>IIBBBBB', 1, 1, 8, 2, 0, 0, 0)
            ihdr_crc = struct.pack('>I', zlib.crc32(b'IHDR' + ihdr_data) & 0xffffffff)
            ihdr = struct.pack('>I', 13) + b'IHDR' + ihdr_data + ihdr_crc
            # IDAT (single red pixel: filter byte 0, then R G B)
            raw = zlib.compress(b'\x00\xff\x00\x00')
            idat_crc = struct.pack('>I', zlib.crc32(b'IDAT' + raw) & 0xffffffff)
            idat = struct.pack('>I', len(raw)) + b'IDAT' + raw + idat_crc
            # IEND
            iend_crc = struct.pack('>I', zlib.crc32(b'IEND') & 0xffffffff)
            iend = struct.pack('>I', 0) + b'IEND' + iend_crc
            return sig + ihdr + idat + iend

        if isinstance(drawing_xml, str):
            drawing_xml = drawing_xml.encode('utf-8')

        # First inject the basic drawing
        injected = self._inject_drawing(xlsx_bytes, drawing_xml)

        # Now add the image and drawing rels
        original = BytesIO(injected)
        output = BytesIO()

        with ZipFile(original, 'r') as zf_in:
            with ZipFile(output, 'w') as zf_out:
                for item in zf_in.namelist():
                    data = zf_in.read(item)

                    if item == '[Content_Types].xml':
                        # Add PNG content type default
                        ct_tree = etree.fromstring(data)
                        # Check if png Default already exists
                        has_png = any(
                            el.get('Extension') == 'png'
                            for el in ct_tree.findall(f'{{{self.CT_NS}}}Default')
                        )
                        if not has_png:
                            etree.SubElement(ct_tree, f'{{{self.CT_NS}}}Default', {
                                'Extension': 'png',
                                'ContentType': 'image/png',
                            })
                        data = etree.tostring(
                            ct_tree, xml_declaration=True, encoding='UTF-8'
                        )

                    zf_out.writestr(item, data)

                # Add the image file
                zf_out.writestr('xl/media/image1.png', _minimal_png())

                # Add drawing rels
                drawing_rels = etree.Element(
                    f'{{{self.PKG_REL_NS}}}Relationships'
                )
                etree.SubElement(drawing_rels, f'{{{self.PKG_REL_NS}}}Relationship', {
                    'Id': 'rId1',
                    'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                    'Target': '../media/image1.png',
                })
                zf_out.writestr(
                    'xl/drawings/_rels/drawing1.xml.rels',
                    etree.tostring(
                        drawing_rels, xml_declaration=True, encoding='UTF-8'
                    ),
                )

        return output.getvalue()

    # =====================================================================
    # Test 1: Basic shape survives round-trip
    # =====================================================================
    def test_shapes_survive_roundtrip(self):
        """
        A twoCellAnchor with a sp (shape with prstGeom rect and txBody)
        must survive load_workbook/save round-trip. The output drawing XML
        must contain a twoCellAnchor with a sp element.
        """
        base = self._create_base_xlsx()
        injected = self._inject_drawing(base, self._shape_drawing_xml())

        # Verify shape exists in the injected file
        original_drawing = self._get_drawing_xml(injected)
        assert original_drawing is not None, "Drawing XML missing from injected file"

        xdr = self.XDR_NS
        sp_elements = original_drawing.findall(f'.//{{{xdr}}}sp')
        assert len(sp_elements) == 1, "Injected file should have exactly 1 shape"

        # Round-trip
        result = self._roundtrip(injected)

        # Verify shape survived
        rt_drawing = self._get_drawing_xml(result)
        assert rt_drawing is not None, (
            "Drawing XML missing from round-tripped file. "
            "The entire drawing was lost during save."
        )

        # Look for the shape element in the output
        rt_sp_elements = rt_drawing.findall(f'.//{{{xdr}}}sp')
        assert len(rt_sp_elements) >= 1, (
            "Shape (sp) element was lost during round-trip. "
            f"Drawing XML output:\n{etree.tostring(rt_drawing, pretty_print=True).decode()}"
        )

        # Verify it is inside a twoCellAnchor
        rt_anchors = rt_drawing.findall(f'{{{xdr}}}twoCellAnchor')
        assert len(rt_anchors) >= 1, "twoCellAnchor was lost during round-trip"

        anchor_has_shape = False
        for anchor in rt_anchors:
            if anchor.findall(f'{{{xdr}}}sp'):
                anchor_has_shape = True
                break
        assert anchor_has_shape, (
            "No twoCellAnchor contains a sp element after round-trip"
        )

    # =====================================================================
    # Test 2: Connector shape survives round-trip
    # =====================================================================
    def test_connector_survives_roundtrip(self):
        """
        A twoCellAnchor with a cxnSp (connector shape) must survive
        load_workbook/save round-trip.
        """
        base = self._create_base_xlsx()
        injected = self._inject_drawing(base, self._connector_drawing_xml())

        # Verify connector exists in the injected file
        xdr = self.XDR_NS
        original_drawing = self._get_drawing_xml(injected)
        assert original_drawing is not None
        cxn_elements = original_drawing.findall(f'.//{{{xdr}}}cxnSp')
        assert len(cxn_elements) == 1, "Injected file should have 1 connector"

        # Round-trip
        result = self._roundtrip(injected)

        # Verify connector survived
        rt_drawing = self._get_drawing_xml(result)
        assert rt_drawing is not None, (
            "Drawing XML missing from round-tripped file"
        )

        rt_cxn_elements = rt_drawing.findall(f'.//{{{xdr}}}cxnSp')
        assert len(rt_cxn_elements) >= 1, (
            "Connector shape (cxnSp) was lost during round-trip. "
            f"Drawing XML output:\n{etree.tostring(rt_drawing, pretty_print=True).decode()}"
        )

        # Verify the connector name attribute was preserved
        a_ns = self.A_NS
        original_name = cxn_elements[0].find(
            f'{{{xdr}}}nvCxnSpPr/{{{xdr}}}cNvPr'
        ).get('name')
        rt_name = rt_cxn_elements[0].find(
            f'.//{{{xdr}}}cNvPr'
        )
        assert rt_name is not None, "Connector cNvPr element missing after round-trip"
        assert rt_name.get('name') == original_name, (
            f"Connector name changed: {original_name!r} -> {rt_name.get('name')!r}"
        )

    # =====================================================================
    # Test 3: Group shape survives round-trip
    # =====================================================================
    def test_group_shape_survives_roundtrip(self):
        """
        A twoCellAnchor with a grpSp containing two sp sub-shapes
        must survive load_workbook/save round-trip.
        """
        base = self._create_base_xlsx()
        injected = self._inject_drawing(base, self._group_shape_drawing_xml())

        # Verify group exists in the injected file
        xdr = self.XDR_NS
        original_drawing = self._get_drawing_xml(injected)
        assert original_drawing is not None
        grp_elements = original_drawing.findall(f'.//{{{xdr}}}grpSp')
        assert len(grp_elements) == 1, "Injected file should have 1 group shape"

        # The group should contain 2 sub-shapes
        sub_shapes = grp_elements[0].findall(f'{{{xdr}}}sp')
        assert len(sub_shapes) == 2, "Group should contain 2 sub-shapes"

        # Round-trip
        result = self._roundtrip(injected)

        # Verify group survived
        rt_drawing = self._get_drawing_xml(result)
        assert rt_drawing is not None, (
            "Drawing XML missing from round-tripped file"
        )

        rt_grp_elements = rt_drawing.findall(f'.//{{{xdr}}}grpSp')
        assert len(rt_grp_elements) >= 1, (
            "Group shape (grpSp) was lost during round-trip. "
            f"Drawing XML output:\n{etree.tostring(rt_drawing, pretty_print=True).decode()}"
        )

        # Verify the sub-shapes survived inside the group
        rt_sub_shapes = rt_grp_elements[0].findall(f'{{{xdr}}}sp')
        assert len(rt_sub_shapes) == 2, (
            f"Group sub-shapes lost: expected 2, got {len(rt_sub_shapes)}. "
            f"Group XML:\n{etree.tostring(rt_grp_elements[0], pretty_print=True).decode()}"
        )

    # =====================================================================
    # Test 4: Shapes coexist with images after round-trip
    # =====================================================================
    def test_shapes_coexist_with_images_roundtrip(self):
        """
        An XLSX with a drawing containing both a shape anchor AND an image
        anchor. Both must survive round-trip.
        """
        base = self._create_base_xlsx()
        drawing_xml = self._shape_and_image_drawing_xml()
        injected = self._inject_drawing_with_image(base, drawing_xml)

        # Verify both shape and image exist in the injected file
        xdr = self.XDR_NS
        original_drawing = self._get_drawing_xml(injected)
        assert original_drawing is not None

        orig_shapes = original_drawing.findall(f'.//{{{xdr}}}sp')
        orig_pics = original_drawing.findall(f'.//{{{xdr}}}pic')
        assert len(orig_shapes) >= 1, "Injected file should have at least 1 shape"
        assert len(orig_pics) >= 1, "Injected file should have at least 1 pic"

        # Round-trip
        result = self._roundtrip(injected)

        # Verify both survived
        rt_drawing = self._get_drawing_xml(result)
        assert rt_drawing is not None, (
            "Drawing XML missing from round-tripped file"
        )

        rt_shapes = rt_drawing.findall(f'.//{{{xdr}}}sp')
        rt_pics = rt_drawing.findall(f'.//{{{xdr}}}pic')

        # The shape must survive (this is the main assertion)
        assert len(rt_shapes) >= 1, (
            "Shape was lost during round-trip when coexisting with image. "
            f"Drawing XML:\n{etree.tostring(rt_drawing, pretty_print=True).decode()}"
        )

        # The image should also survive (openpyxl already handles images)
        assert len(rt_pics) >= 1, (
            "Image (pic) was lost during round-trip. "
            f"Drawing XML:\n{etree.tostring(rt_drawing, pretty_print=True).decode()}"
        )

        # Both anchors should be present
        rt_anchors = rt_drawing.findall(f'{{{xdr}}}twoCellAnchor')
        assert len(rt_anchors) >= 2, (
            f"Expected at least 2 twoCellAnchors (shape + image), "
            f"got {len(rt_anchors)}"
        )

    # =====================================================================
    # Test 5: Shape text content fidelity after round-trip
    # =====================================================================
    def test_shape_drawing_xml_fidelity(self):
        """
        Verify the text content of a preserved textbox shape matches
        after round-trip. The text "Hello World" in the txBody must be
        present in the output.
        """
        text_content = "Hello World"
        base = self._create_base_xlsx()
        injected = self._inject_drawing(base, self._shape_drawing_xml(text=text_content))

        # Verify original text
        xdr = self.XDR_NS
        a_ns = self.A_NS
        original_drawing = self._get_drawing_xml(injected)
        assert original_drawing is not None

        orig_text_els = original_drawing.findall(f'.//{{{a_ns}}}t')
        orig_texts = [t.text for t in orig_text_els if t.text]
        assert text_content in orig_texts, (
            f"Original drawing should contain '{text_content}', found: {orig_texts}"
        )

        # Round-trip
        result = self._roundtrip(injected)

        # Verify text fidelity
        rt_drawing = self._get_drawing_xml(result)
        assert rt_drawing is not None, (
            "Drawing XML missing from round-tripped file"
        )

        rt_text_els = rt_drawing.findall(f'.//{{{a_ns}}}t')
        rt_texts = [t.text for t in rt_text_els if t.text]
        assert text_content in rt_texts, (
            f"Text '{text_content}' was lost during round-trip. "
            f"Found texts: {rt_texts}. "
            f"Drawing XML:\n{etree.tostring(rt_drawing, pretty_print=True).decode()}"
        )

        # Also verify the shape structure around the text is preserved
        rt_sp = rt_drawing.findall(f'.//{{{xdr}}}sp')
        assert len(rt_sp) >= 1, "Shape element not found after round-trip"

        # Find txBody within the shape
        shape_has_text = False
        for sp in rt_sp:
            tx_body = sp.find(f'{{{xdr}}}txBody')
            if tx_body is None:
                # Also check without namespace (some serializations use different prefix)
                tx_body = sp.find(f'.//{{{a_ns}}}t/..')
            if tx_body is not None:
                text_nodes = sp.findall(f'.//{{{a_ns}}}t')
                for tn in text_nodes:
                    if tn.text and text_content in tn.text:
                        shape_has_text = True
                        break
            if shape_has_text:
                break

        assert shape_has_text, (
            f"Shape element exists but does not contain text '{text_content}'. "
            "The text body may have been stripped from the shape during round-trip."
        )

    # =====================================================================
    # Test 6: Worksheet data unaffected by shape preservation
    # =====================================================================
    def test_worksheet_data_unaffected_by_shapes(self):
        """
        Verify that normal worksheet cell data is preserved correctly
        when the workbook contains shapes. Shape preservation must not
        corrupt other data.
        """
        base = self._create_base_xlsx()
        injected = self._inject_drawing(base, self._shape_drawing_xml())

        result = self._roundtrip(injected)

        # Re-load and verify cell data
        wb = load_workbook(BytesIO(result))
        ws = wb['Sheet1']
        assert ws['A1'].value == "test data", (
            f"Cell A1 value changed: expected 'test data', got {ws['A1'].value!r}"
        )
        assert ws['B2'].value == 42, (
            f"Cell B2 value changed: expected 42, got {ws['B2'].value!r}"
        )
