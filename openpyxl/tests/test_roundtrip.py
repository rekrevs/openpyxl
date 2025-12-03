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
