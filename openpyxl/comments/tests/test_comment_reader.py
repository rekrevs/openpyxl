# Copyright (c) 2010-2024 openpyxl

from openpyxl.reader.excel import load_workbook
from openpyxl.xml.functions import fromstring

from ..comments import Comment

import pytest


def test_read_comments(datadir):
    datadir.chdir()
    from .. comment_sheet import CommentSheet

    with open("comments2.xml") as src:
        node = fromstring(src.read())

    sheet = CommentSheet.from_tree(node)
    comments = list(sheet.comments)
    assert comments == [
        ('A1', Comment('Cuke:\nFirst Comment', 'Cuke')),
        ('D1', Comment('Cuke:\nSecond Comment', 'Cuke')),
        ('A2', Comment('Not Cuke:\nThird Comment', 'Not Cuke'))
         ]


def test_comments_cell_association(datadir):
    datadir.chdir()
    wb = load_workbook('comments.xlsx')
    assert wb['Sheet1']["A1"].comment.author == "Cuke"
    assert wb['Sheet1']["A1"].comment.text == "Cuke:\nFirst Comment"
    assert wb['Sheet2']["A1"].comment is None
    assert wb['Sheet1']["D1"].comment.text == "Cuke:\nSecond Comment"


def test_comments_with_iterators(datadir):
    datadir.chdir()
    wb = load_workbook('comments.xlsx', read_only=True)
    ws = wb['Sheet1']
    with pytest.raises(AttributeError):
        assert ws["A1"].comment.author == "Cuke"


def test_comment_formatting_preserved(datadir):
    """Test that rich text formatting is preserved on comments"""
    datadir.chdir()
    from ..comment_sheet import CommentSheet

    with open("comments2.xml") as src:
        node = fromstring(src.read())

    sheet = CommentSheet.from_tree(node)
    comments = list(sheet.comments)

    # Check that rich text object is preserved
    ref, comment = comments[0]
    assert ref == 'A1'
    assert comment._text_obj is not None
    assert len(comment._text_obj.r) == 2  # Two rich text runs

    # First run should have bold formatting
    first_run = comment._text_obj.r[0]
    assert first_run.rPr is not None
    assert first_run.rPr.b is True  # Bold

    # Plain text content should still work
    assert comment.text == "Cuke:\nFirst Comment"


def test_comment_formatting_roundtrip(datadir):
    """Test that rich text formatting survives round-trip"""
    datadir.chdir()
    from ..comment_sheet import CommentSheet, CommentRecord

    with open("comments2.xml") as src:
        node = fromstring(src.read())

    sheet = CommentSheet.from_tree(node)
    comments = list(sheet.comments)

    # Simulate assigning comment to a cell and writing it back
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active

    ref, comment = comments[0]
    ws['A1'].comment = comment

    # Convert back to CommentRecord (what happens on save)
    record = CommentRecord.from_cell(ws['A1'])

    # Verify rich text is preserved
    assert record.text is not None
    assert len(record.text.r) == 2  # Two rich text runs preserved
    assert record.text.r[0].rPr.b is True  # Bold formatting preserved
