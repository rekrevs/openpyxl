# Copyright (c) 2010-2024 openpyxl

"""
Support for sparkline mini-charts in cells.

Sparklines are stored in the extLst element of the worksheet with
GUID: {05C60535-1F16-4FD2-B633-F4F36F0B64E0}

Namespace: http://schemas.microsoft.com/office/spreadsheetml/2009/9/main (x14)
"""

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    String,
    Bool,
    Integer,
    Float,
    Typed,
    NoneSet,
)
from openpyxl.descriptors.nested import NestedText
from openpyxl.descriptors.sequence import Sequence, NestedSequence

# Sparkline extension GUID
SPARKLINE_GUID = "{05C60535-1F16-4FD2-B633-F4F36F0B64E0}"

# x14 namespace
X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"


class SparklineColor(Serialisable):
    """Color specification for sparkline elements"""

    tagname = "color"

    rgb = String(allow_none=True)
    theme = Integer(allow_none=True)
    tint = Float(allow_none=True)

    def __init__(self, rgb=None, theme=None, tint=None):
        self.rgb = rgb
        self.theme = theme
        self.tint = tint


class Sparkline(Serialisable):
    """
    Individual sparkline definition.

    Attributes
    ----------
    f : str
        Data range formula (e.g., "Sheet1!A1:A10")
    sqref : str
        Cell where sparkline is displayed (e.g., "B1")
    """

    tagname = "sparkline"
    namespace = X14_NS

    f = NestedText(expected_type=str, allow_none=True)  # Data range formula
    sqref = NestedText(expected_type=str, allow_none=True)  # Display cell

    __elements__ = ('f', 'sqref')

    def __init__(self, f=None, sqref=None):
        self.f = f
        self.sqref = sqref


class SparklineGroup(Serialisable):
    """
    Group of sparklines sharing common properties.

    Attributes
    ----------
    type : str
        Sparkline type: 'line', 'column', or 'stacked' (win/loss)
    displayEmptyCellsAs : str
        How to display empty cells: 'gap', 'zero', or 'span'
    """

    tagname = "sparklineGroup"
    namespace = X14_NS

    # Sparkline type
    type = NoneSet(values=['line', 'column', 'stacked'])

    # Display options
    displayEmptyCellsAs = NoneSet(values=['gap', 'zero', 'span'])
    markers = Bool(allow_none=True)
    high = Bool(allow_none=True)
    low = Bool(allow_none=True)
    first = Bool(allow_none=True)
    last = Bool(allow_none=True)
    negative = Bool(allow_none=True)
    displayXAxis = Bool(allow_none=True)
    displayHidden = Bool(allow_none=True)
    rightToLeft = Bool(allow_none=True)

    # Min/max axis
    minAxisType = NoneSet(values=['individual', 'group', 'custom'])
    maxAxisType = NoneSet(values=['individual', 'group', 'custom'])
    manualMin = Float(allow_none=True)
    manualMax = Float(allow_none=True)

    # Line weight
    lineWeight = Float(allow_none=True)

    # Date axis
    dateAxis = Bool(allow_none=True)

    # Colors
    colorSeries = Typed(expected_type=SparklineColor, allow_none=True)
    colorNegative = Typed(expected_type=SparklineColor, allow_none=True)
    colorAxis = Typed(expected_type=SparklineColor, allow_none=True)
    colorMarkers = Typed(expected_type=SparklineColor, allow_none=True)
    colorFirst = Typed(expected_type=SparklineColor, allow_none=True)
    colorLast = Typed(expected_type=SparklineColor, allow_none=True)
    colorHigh = Typed(expected_type=SparklineColor, allow_none=True)
    colorLow = Typed(expected_type=SparklineColor, allow_none=True)

    # Sparklines in this group (wrapped in <sparklines> element)
    sparklines = NestedSequence(expected_type=Sparkline)

    __elements__ = ('colorSeries', 'colorNegative', 'colorAxis', 'colorMarkers',
                    'colorFirst', 'colorLast', 'colorHigh', 'colorLow', 'sparklines')

    def __init__(
        self,
        type=None,
        displayEmptyCellsAs=None,
        markers=None,
        high=None,
        low=None,
        first=None,
        last=None,
        negative=None,
        displayXAxis=None,
        displayHidden=None,
        rightToLeft=None,
        minAxisType=None,
        maxAxisType=None,
        manualMin=None,
        manualMax=None,
        lineWeight=None,
        dateAxis=None,
        colorSeries=None,
        colorNegative=None,
        colorAxis=None,
        colorMarkers=None,
        colorFirst=None,
        colorLast=None,
        colorHigh=None,
        colorLow=None,
        sparklines=(),
    ):
        self.type = type
        self.displayEmptyCellsAs = displayEmptyCellsAs
        self.markers = markers
        self.high = high
        self.low = low
        self.first = first
        self.last = last
        self.negative = negative
        self.displayXAxis = displayXAxis
        self.displayHidden = displayHidden
        self.rightToLeft = rightToLeft
        self.minAxisType = minAxisType
        self.maxAxisType = maxAxisType
        self.manualMin = manualMin
        self.manualMax = manualMax
        self.lineWeight = lineWeight
        self.dateAxis = dateAxis
        self.colorSeries = colorSeries
        self.colorNegative = colorNegative
        self.colorAxis = colorAxis
        self.colorMarkers = colorMarkers
        self.colorFirst = colorFirst
        self.colorLast = colorLast
        self.colorHigh = colorHigh
        self.colorLow = colorLow
        self.sparklines = sparklines


class SparklineGroups(Serialisable):
    """
    Container for all sparkline groups in a worksheet.
    """

    tagname = "sparklineGroups"
    namespace = X14_NS

    sparklineGroup = Sequence(expected_type=SparklineGroup)

    __elements__ = ('sparklineGroup',)

    def __init__(self, sparklineGroup=()):
        self.sparklineGroup = sparklineGroup

    def __len__(self):
        return len(self.sparklineGroup)

    def __bool__(self):
        return bool(self.sparklineGroup)

    def append(self, group):
        """Add a sparkline group"""
        self.sparklineGroup.append(group)
