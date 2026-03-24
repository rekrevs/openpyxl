
# Copyright (c) 2010-2024 openpyxl


from copy import deepcopy
from io import BytesIO
from warnings import warn

from openpyxl.xml.functions import fromstring
from openpyxl.xml.constants import IMAGE_NS
from openpyxl.packaging.relationship import (
    get_rel,
    get_rels_path,
    get_dependents,
)
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl.drawing.image import Image, PILImage
from openpyxl.chart.chartspace import ChartSpace
from openpyxl.chart.reader import read_chart


# Local names of elements that indicate a shape anchor (not chart/image)
_SHAPE_CHILD_TAGS = frozenset({"sp", "cxnSp", "grpSp"})

# Local names of anchor elements in SpreadsheetDrawing
_ANCHOR_TAGS = frozenset({"twoCellAnchor", "oneCellAnchor", "absoluteAnchor"})


def _collect_shape_anchors(tree, chart_rel_ids):
    """
    Walk the raw drawing XML tree and collect (as deep copies) any anchor
    elements that contain shapes rather than charts or images.

    An anchor is considered a "shape anchor" if it contains a child with
    local name in _SHAPE_CHILD_TAGS, OR if it contains a graphicFrame
    whose relationship id is NOT in chart_rel_ids (i.e. SmartArt, etc.).

    An anchor that contains only `pic` or a chart graphicFrame is already
    handled by the existing chart/image extraction and is skipped here.

    Parameters
    ----------
    tree : lxml.etree._Element
        The root <xdr:wsDr> element of the drawing XML.
    chart_rel_ids : set[str]
        Set of relationship IDs that point to charts (from _chart_rels).

    Returns
    -------
    list[lxml.etree._Element]
        Deep-copied anchor elements containing shapes.
    """
    shape_anchors = []

    for child in tree:
        local = child.tag.rpartition("}")[-1] if "}" in child.tag else child.tag
        if local not in _ANCHOR_TAGS:
            continue

        # Check what kind of content this anchor has
        has_shape_content = False

        for sub in child:
            sub_local = sub.tag.rpartition("}")[-1] if "}" in sub.tag else sub.tag

            if sub_local in _SHAPE_CHILD_TAGS:
                has_shape_content = True
                break

            if sub_local == "graphicFrame":
                # Check if this graphicFrame references a chart
                # Look for r:id in the graphic -> graphicData -> chart element
                is_chart = False
                for desc in sub.iter():
                    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    rid = desc.get(f"{{{rel_ns}}}id")
                    if rid and rid in chart_rel_ids:
                        is_chart = True
                        break
                if not is_chart:
                    # graphicFrame without chart reference = SmartArt or similar
                    has_shape_content = True
                    break

            if sub_local == "pic":
                # Image anchor — already handled
                pass

        if has_shape_content:
            shape_anchors.append(deepcopy(child))

    return shape_anchors


def find_images(archive, path):
    """
    Given the path to a drawing file extract charts, images, and shape anchors.

    Ignore errors due to unsupported parts of DrawingML.

    Returns
    -------
    tuple of (charts, images, shape_anchors)
        charts: list of Chart objects
        images: list of Image objects
        shape_anchors: list of raw XML anchor elements containing shapes
    """

    src = archive.read(path)
    tree = fromstring(src)
    try:
        drawing = SpreadsheetDrawing.from_tree(tree)
    except TypeError:
        warn("DrawingML support is incomplete and limited to charts and images only. Shapes and drawings will be lost.")
        return [], [], []

    rels_path = get_rels_path(path)
    deps = []
    if rels_path in archive.namelist():
        deps = get_dependents(archive, rels_path)

    charts = []
    chart_rel_ids = set()
    for rel in drawing._chart_rels:
        try:
            cs = get_rel(archive, deps, rel.id, ChartSpace)
        except TypeError as e:
            warn(f"Unable to read chart {rel.id} from {path} {e}")
            continue
        chart = read_chart(cs)
        chart.anchor = rel.anchor
        charts.append(chart)
        chart_rel_ids.add(rel.id)

    images = []
    if not PILImage: # Pillow not installed, drop images
        shape_anchors = _collect_shape_anchors(tree, chart_rel_ids)
        return charts, images, shape_anchors

    for rel in drawing._blip_rels:
        dep = deps.get(rel.embed)
        if dep.Type == IMAGE_NS:
            try:
                image = Image(BytesIO(archive.read(dep.target)))
            except OSError:
                msg = "The image {0} will be removed because it cannot be read".format(dep.target)
                warn(msg)
                continue
            if image.format.upper() == "WMF": # cannot save
                msg = "{0} image format is not supported so the image is being dropped".format(image.format)
                warn(msg)
                continue
            image.anchor = rel.anchor
            images.append(image)

    # Collect shape anchors from raw XML (additive, does not interfere with charts/images)
    shape_anchors = _collect_shape_anchors(tree, chart_rel_ids)

    return charts, images, shape_anchors
