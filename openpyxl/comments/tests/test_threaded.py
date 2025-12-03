# Copyright (c) 2010-2024 openpyxl

import pytest
from datetime import datetime

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def Person():
    from ..person import Person
    return Person


@pytest.fixture
def PersonList():
    from ..person import PersonList
    return PersonList


@pytest.fixture
def ThreadedComment():
    from ..threaded import ThreadedComment
    return ThreadedComment


@pytest.fixture
def ThreadedCommentList():
    from ..threaded import ThreadedCommentList
    return ThreadedCommentList


@pytest.fixture
def Mention():
    from ..threaded import Mention
    return Mention


class TestPerson:

    def test_ctor(self, Person):
        person = Person(
            displayName="John Doe",
            id="{12345678-1234-5678-9ABC-DEF012345678}",
            userId="johnd",
            providerId="provider123"
        )
        assert person.displayName == "John Doe"
        assert person.id == "{12345678-1234-5678-9ABC-DEF012345678}"
        assert person.userId == "johnd"
        assert person.providerId == "provider123"

    def test_from_xml(self, Person):
        src = """
        <person xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/revision16"
                displayName="Jane Smith"
                id="{ABCD1234-5678-9ABC-DEF0-123456789ABC}"
                userId="janes"
                providerId="Windows Live"/>
        """
        node = fromstring(src)
        person = Person.from_tree(node)
        assert person.displayName == "Jane Smith"
        assert person.id == "{ABCD1234-5678-9ABC-DEF0-123456789ABC}"


class TestPersonList:

    def test_ctor(self, PersonList, Person):
        persons = PersonList(
            person=[
                Person(displayName="User 1", id="{11111111-1111-1111-1111-111111111111}"),
                Person(displayName="User 2", id="{22222222-2222-2222-2222-222222222222}"),
            ]
        )
        assert len(persons) == 2

    def test_get_by_id(self, PersonList, Person):
        persons = PersonList(
            person=[
                Person(displayName="User 1", id="{11111111-1111-1111-1111-111111111111}"),
                Person(displayName="User 2", id="{22222222-2222-2222-2222-222222222222}"),
            ]
        )
        user1 = persons["{11111111-1111-1111-1111-111111111111}"]
        assert user1.displayName == "User 1"

    def test_get_missing(self, PersonList, Person):
        persons = PersonList()
        with pytest.raises(KeyError):
            persons["{00000000-0000-0000-0000-000000000000}"]

    def test_from_xml(self, PersonList):
        src = """
        <personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/revision16">
            <person displayName="Alice" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"/>
            <person displayName="Bob" id="{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}"/>
        </personList>
        """
        node = fromstring(src)
        persons = PersonList.from_tree(node)
        assert len(persons) == 2


class TestThreadedComment:

    def test_ctor(self, ThreadedComment):
        comment = ThreadedComment(
            ref="A1",
            dT=datetime(2024, 1, 15, 10, 30, 0),
            personId="{11111111-1111-1111-1111-111111111111}",
            id="{22222222-2222-2222-2222-222222222222}",
            text="This is a comment"
        )
        assert comment.ref == "A1"
        assert comment.text == "This is a comment"
        assert not comment.is_reply

    def test_reply(self, ThreadedComment):
        reply = ThreadedComment(
            ref="A1",
            personId="{11111111-1111-1111-1111-111111111111}",
            id="{33333333-3333-3333-3333-333333333333}",
            parentId="{22222222-2222-2222-2222-222222222222}",
            text="This is a reply"
        )
        assert reply.is_reply
        assert reply.parentId == "{22222222-2222-2222-2222-222222222222}"

    def test_resolved(self, ThreadedComment):
        comment = ThreadedComment(
            ref="A1",
            personId="{11111111-1111-1111-1111-111111111111}",
            id="{22222222-2222-2222-2222-222222222222}",
            done=True,
            text="Resolved comment"
        )
        assert comment.is_resolved

    def test_from_xml(self, ThreadedComment):
        src = """
        <threadedComment xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
                         ref="B2"
                         dT="2024-03-15T14:30:00.00Z"
                         personId="{12345678-1234-5678-9ABC-DEF012345678}"
                         id="{ABCDEF12-3456-7890-ABCD-EF1234567890}">
            <text>Hello, this is a test comment</text>
        </threadedComment>
        """
        node = fromstring(src)
        comment = ThreadedComment.from_tree(node)
        assert comment.ref == "B2"
        assert comment.text == "Hello, this is a test comment"
        assert comment.personId == "{12345678-1234-5678-9ABC-DEF012345678}"


class TestThreadedCommentList:

    def test_ctor(self, ThreadedCommentList, ThreadedComment):
        comments = ThreadedCommentList(
            threadedComment=[
                ThreadedComment(ref="A1", id="{11111111-1111-1111-1111-111111111111}", personId="{00000000-0000-0000-0000-000000000000}", text="Comment 1"),
                ThreadedComment(ref="A1", id="{22222222-2222-2222-2222-222222222222}", personId="{00000000-0000-0000-0000-000000000000}", parentId="{11111111-1111-1111-1111-111111111111}", text="Reply to 1"),
            ]
        )
        assert len(comments) == 2

    def test_get_by_ref(self, ThreadedCommentList, ThreadedComment):
        comments = ThreadedCommentList(
            threadedComment=[
                ThreadedComment(ref="A1", id="{11111111-1111-1111-1111-111111111111}", personId="{00000000-0000-0000-0000-000000000000}", text="A1 comment"),
                ThreadedComment(ref="B2", id="{22222222-2222-2222-2222-222222222222}", personId="{00000000-0000-0000-0000-000000000000}", text="B2 comment"),
                ThreadedComment(ref="A1", id="{33333333-3333-3333-3333-333333333333}", personId="{00000000-0000-0000-0000-000000000000}", parentId="{11111111-1111-1111-1111-111111111111}", text="A1 reply"),
            ]
        )
        a1_comments = comments.get_by_ref("A1")
        assert len(a1_comments) == 2

    def test_get_thread(self, ThreadedCommentList, ThreadedComment):
        dt = datetime(2024, 1, 1, 10, 0, 0)
        comments = ThreadedCommentList(
            threadedComment=[
                ThreadedComment(ref="A1", id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}", personId="{00000000-0000-0000-0000-000000000000}", dT=dt, text="Root"),
                ThreadedComment(ref="A1", id="{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}", personId="{00000000-0000-0000-0000-000000000000}", parentId="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}", dT=dt, text="Reply 1"),
                ThreadedComment(ref="A1", id="{CCCCCCCC-CCCC-CCCC-CCCC-CCCCCCCCCCCC}", personId="{00000000-0000-0000-0000-000000000000}", parentId="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}", dT=dt, text="Reply 2"),
            ]
        )
        thread = comments.get_thread("{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}")
        assert len(thread) == 3
        assert thread[0].id == "{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"

    def test_from_xml(self, ThreadedCommentList):
        src = """
        <ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
            <threadedComment ref="A1" id="{11111111-1111-1111-1111-111111111111}" personId="{00000000-0000-0000-0000-000000000000}" dT="2024-01-01T00:00:00Z">
                <text>First comment</text>
            </threadedComment>
            <threadedComment ref="A1" id="{22222222-2222-2222-2222-222222222222}" personId="{00000000-0000-0000-0000-000000000000}" parentId="{11111111-1111-1111-1111-111111111111}" dT="2024-01-01T00:01:00Z">
                <text>Reply</text>
            </threadedComment>
        </ThreadedComments>
        """
        node = fromstring(src)
        comments = ThreadedCommentList.from_tree(node)
        assert len(comments) == 2
        assert comments.threadedComment[0].text == "First comment"
        assert comments.threadedComment[1].parentId == "{11111111-1111-1111-1111-111111111111}"

    def test_round_trip(self, ThreadedCommentList, ThreadedComment):
        dt = datetime(2024, 1, 15, 10, 30, 0)
        original = ThreadedCommentList(
            threadedComment=[
                ThreadedComment(
                    ref="C3",
                    dT=dt,
                    personId="{00000000-0000-0000-0000-000000000000}",
                    id="{11111111-1111-1111-1111-111111111111}",
                    text="Test round trip"
                )
            ]
        )

        # Serialize
        tree = original.to_tree()
        xml = tostring(tree)

        # Deserialize
        parsed = ThreadedCommentList.from_tree(fromstring(xml))

        # Verify
        assert len(parsed) == 1
        assert parsed.threadedComment[0].ref == "C3"
        assert parsed.threadedComment[0].text == "Test round trip"


class TestWorkbookIntegration:
    """Integration tests for threaded comments with workbook round-trip"""

    def test_threaded_comments_workbook_round_trip(self, PersonList, Person, ThreadedCommentList, ThreadedComment):
        """Test that threaded comments survive a workbook save/load cycle"""
        from io import BytesIO
        from openpyxl import Workbook, load_workbook

        # Create workbook with threaded comments
        wb = Workbook()
        ws = wb.active
        ws['A1'] = "Test"

        # Create person list
        wb.persons = PersonList(person=[
            Person(
                displayName="Test User",
                id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}",
                userId="testuser"
            )
        ])

        # Create threaded comments
        dt = datetime(2024, 6, 15, 14, 30, 0)
        ws.threaded_comments = ThreadedCommentList(threadedComment=[
            ThreadedComment(
                ref="A1",
                dT=dt,
                personId="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}",
                id="{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}",
                text="This is a test comment"
            ),
            ThreadedComment(
                ref="A1",
                dT=dt,
                personId="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}",
                id="{CCCCCCCC-CCCC-CCCC-CCCC-CCCCCCCCCCCC}",
                parentId="{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}",
                text="This is a reply"
            )
        ])

        # Save and reload
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        wb2 = load_workbook(buffer)
        ws2 = wb2.active

        # Verify persons
        assert wb2.persons is not None
        assert len(wb2.persons) == 1
        assert wb2.persons.person[0].displayName == "Test User"

        # Verify threaded comments
        assert ws2.threaded_comments is not None
        assert len(ws2.threaded_comments) == 2
        assert ws2.threaded_comments.threadedComment[0].text == "This is a test comment"
        assert ws2.threaded_comments.threadedComment[1].text == "This is a reply"
        assert ws2.threaded_comments.threadedComment[1].parentId == "{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}"
