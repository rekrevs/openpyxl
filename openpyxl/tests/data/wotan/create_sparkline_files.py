# Copyright (c) 2010-2024 openpyxl
"""
Script to create sparkline test files programmatically.

Sparklines are written via extensions, so we need to add them to the worksheet
extensions with the proper GUID.
"""

from openpyxl import Workbook
from openpyxl.worksheet.sparkline import (
    Sparkline, SparklineGroup, SparklineGroups, SparklineColor, SPARKLINE_GUID
)
from openpyxl.descriptors.excel import ExtensionList, Extension
from lxml.etree import Element, SubElement, tostring, QName

X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM_NS = "http://schemas.microsoft.com/office/excel/2006/main"

NSMAP = {
    'x14': X14_NS,
    'xm': XM_NS,
}


def create_sparkline_extension(sparkline_groups):
    """Create an Extension element with sparkline groups."""
    # Create the ext element with proper namespace declarations
    ext = Element('ext', uri=SPARKLINE_GUID, nsmap=NSMAP)

    # Add sparkline groups
    groups_el = SubElement(ext, f'{{{X14_NS}}}sparklineGroups')

    for group in sparkline_groups.sparklineGroup:
        group_el = SubElement(groups_el, f'{{{X14_NS}}}sparklineGroup')

        # Set type
        if group.type:
            group_el.set('type', group.type)

        # Set display options
        if group.displayEmptyCellsAs:
            group_el.set('displayEmptyCellsAs', group.displayEmptyCellsAs)
        if group.markers:
            group_el.set('markers', '1')
        if group.high:
            group_el.set('high', '1')
        if group.low:
            group_el.set('low', '1')
        if group.first:
            group_el.set('first', '1')
        if group.last:
            group_el.set('last', '1')
        if group.negative:
            group_el.set('negative', '1')

        # Add colors if specified
        if group.colorSeries:
            color_el = SubElement(group_el, f'{{{X14_NS}}}colorSeries')
            if group.colorSeries.rgb:
                color_el.set('rgb', group.colorSeries.rgb)
            if group.colorSeries.theme is not None:
                color_el.set('theme', str(group.colorSeries.theme))

        if group.colorNegative:
            color_el = SubElement(group_el, f'{{{X14_NS}}}colorNegative')
            if group.colorNegative.rgb:
                color_el.set('rgb', group.colorNegative.rgb)
            if group.colorNegative.theme is not None:
                color_el.set('theme', str(group.colorNegative.theme))

        # Add sparklines
        sparklines_el = SubElement(group_el, f'{{{X14_NS}}}sparklines')
        for sparkline in group.sparklines:
            sp_el = SubElement(sparklines_el, f'{{{X14_NS}}}sparkline')
            f_el = SubElement(sp_el, f'{{{XM_NS}}}f')
            f_el.text = sparkline.f
            sqref_el = SubElement(sp_el, f'{{{XM_NS}}}sqref')
            sqref_el.text = sparkline.sqref

    return Extension.from_tree(ext)


def create_line_sparkline():
    """Create a line sparkline test file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Add data
    data = [10, 25, 15, 30, 20, 35, 25, 40, 30, 45]
    for i, val in enumerate(data, 1):
        ws.cell(row=i, column=1, value=val)

    # Create sparkline
    sparkline = Sparkline(f="Sheet1!A1:A10", sqref="B1")
    group = SparklineGroup(
        type='line',
        displayEmptyCellsAs='gap',
        colorSeries=SparklineColor(theme=4),
        sparklines=[sparkline]
    )
    groups = SparklineGroups(sparklineGroup=[group])

    # Add to extensions
    ext = create_sparkline_extension(groups)
    ws.extensions = ExtensionList(ext=[ext])

    wb.save('sparklines/line-sparkline.xlsx')
    print("Created sparklines/line-sparkline.xlsx")


def create_column_sparkline():
    """Create a column sparkline test file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Add data
    data = [50, 75, 60, 85, 70, 90, 65, 80, 55, 95]
    for i, val in enumerate(data, 1):
        ws.cell(row=i, column=1, value=val)

    # Create sparkline
    sparkline = Sparkline(f="Sheet1!A1:A10", sqref="B1")
    group = SparklineGroup(
        type='column',
        displayEmptyCellsAs='zero',
        colorSeries=SparklineColor(theme=5),
        sparklines=[sparkline]
    )
    groups = SparklineGroups(sparklineGroup=[group])

    # Add to extensions
    ext = create_sparkline_extension(groups)
    ws.extensions = ExtensionList(ext=[ext])

    wb.save('sparklines/column-sparkline.xlsx')
    print("Created sparklines/column-sparkline.xlsx")


def create_winloss_sparkline():
    """Create a win/loss sparkline test file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Add data (positive = win, negative = loss)
    data = [1, -1, 1, 1, -1, 1, -1, -1, 1, 1]
    for i, val in enumerate(data, 1):
        ws.cell(row=i, column=1, value=val)

    # Create sparkline - stacked type for win/loss
    sparkline = Sparkline(f="Sheet1!A1:A10", sqref="B1")
    group = SparklineGroup(
        type='stacked',
        displayEmptyCellsAs='zero',
        negative=True,
        colorSeries=SparklineColor(theme=4),
        colorNegative=SparklineColor(rgb="FFFF0000"),
        sparklines=[sparkline]
    )
    groups = SparklineGroups(sparklineGroup=[group])

    # Add to extensions
    ext = create_sparkline_extension(groups)
    ws.extensions = ExtensionList(ext=[ext])

    wb.save('sparklines/winloss-sparkline.xlsx')
    print("Created sparklines/winloss-sparkline.xlsx")


if __name__ == '__main__':
    import os
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    create_line_sparkline()
    create_column_sparkline()
    create_winloss_sparkline()

    print("\nAll sparkline test files created!")
