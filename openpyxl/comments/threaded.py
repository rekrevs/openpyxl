# Copyright (c) 2010-2024 openpyxl

"""
Support for threaded comments in XLSX files.

Threaded comments (introduced in Excel 2019/Office 365) support:
- Multiple replies in a thread
- @mentions
- Resolution status (done)
- Timestamps

The threaded comments are stored at xl/threadedComments/threadedComment{n}.xml
"""

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import String, Bool, DateTime
from openpyxl.descriptors.excel import Guid
from openpyxl.descriptors.sequence import Sequence
from openpyxl.descriptors.nested import NestedText

from openpyxl.xml.constants import XLTHREADED_NS


class Mention(Serialisable):
    """
    Represents an @mention in a threaded comment.

    Attributes
    ----------
    mentionpersonId : str
        GUID of the mentioned person
    mentionId : str
        Unique ID for this mention
    startIndex : int
        Starting character index of the mention
    length : int
        Length of the mention text
    """

    tagname = "mention"
    namespace = XLTHREADED_NS

    mentionpersonId = Guid()
    mentionId = Guid()
    startIndex = String()
    length = String()

    __attrs__ = ('mentionpersonId', 'mentionId', 'startIndex', 'length')

    def __init__(self, mentionpersonId=None, mentionId=None, startIndex=None, length=None):
        self.mentionpersonId = mentionpersonId
        self.mentionId = mentionId
        self.startIndex = startIndex
        self.length = length


class ThreadedComment(Serialisable):
    """
    Represents a single threaded comment.

    Attributes
    ----------
    ref : str
        Cell reference (e.g., "A1")
    dT : datetime
        Timestamp when the comment was created/modified
    personId : str
        GUID of the person who wrote the comment
    id : str
        GUID uniquely identifying this comment
    parentId : str, optional
        GUID of the parent comment (for replies)
    done : bool, optional
        Whether this comment thread is resolved
    text : str
        The comment text content
    mentions : list
        List of @mentions in the comment
    """

    tagname = "threadedComment"
    namespace = XLTHREADED_NS

    ref = String()
    dT = DateTime(allow_none=True)
    personId = Guid()
    id = Guid()
    parentId = Guid(allow_none=True)
    done = Bool(allow_none=True)

    text = NestedText(expected_type=str, allow_none=True)
    mentions = Sequence(expected_type=Mention)

    __attrs__ = ('ref', 'dT', 'personId', 'id', 'parentId', 'done')
    __elements__ = ('text', 'mentions')

    def __init__(
        self,
        ref=None,
        dT=None,
        personId=None,
        id=None,
        parentId=None,
        done=None,
        text=None,
        mentions=(),
    ):
        self.ref = ref
        self.dT = dT
        self.personId = personId
        self.id = id
        self.parentId = parentId
        self.done = done
        self.text = text
        self.mentions = mentions

    @property
    def is_reply(self):
        """Check if this comment is a reply to another comment"""
        return self.parentId is not None

    @property
    def is_resolved(self):
        """Check if this comment thread is marked as resolved"""
        return self.done is True


class ThreadedCommentList(Serialisable):
    """
    Container for all threaded comments in a worksheet.

    This is stored at xl/threadedComments/threadedComment{n}.xml
    in the XLSX archive.
    """

    tagname = "ThreadedComments"
    namespace = XLTHREADED_NS

    threadedComment = Sequence(expected_type=ThreadedComment)

    _id = None
    _path = "/xl/threadedComments/threadedComment{0}.xml"
    mime_type = "application/vnd.ms-excel.threadedcomments+xml"
    _rel_type = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
    _rel_id = None

    __elements__ = ('threadedComment',)

    def __init__(self, threadedComment=()):
        self.threadedComment = threadedComment

    def __len__(self):
        return len(self.threadedComment)

    def __bool__(self):
        return bool(self.threadedComment)


    def append(self, comment):
        """Add a threaded comment"""
        self.threadedComment.append(comment)

    def get_by_ref(self, ref):
        """Get all threaded comments for a cell reference"""
        return [c for c in self.threadedComment if c.ref == ref]

    def get_thread(self, comment_id):
        """
        Get a comment and all its replies.

        Parameters
        ----------
        comment_id : str
            The GUID of the root comment

        Returns
        -------
        list
            List of ThreadedComment objects in thread order
        """
        thread = []

        # Find the root comment
        root = None
        for c in self.threadedComment:
            if c.id == comment_id:
                root = c
                break

        if root is None:
            return thread

        thread.append(root)

        # Find all replies (recursively)
        def find_replies(parent_id):
            replies = []
            for c in self.threadedComment:
                if c.parentId == parent_id:
                    replies.append(c)
            # Sort by timestamp
            replies.sort(key=lambda x: x.dT if x.dT else "")
            for reply in replies:
                thread.append(reply)
                find_replies(reply.id)

        find_replies(comment_id)
        return thread

    @property
    def path(self):
        """Return path within the archive"""
        return self._path.format(self._id)
