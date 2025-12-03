# Copyright (c) 2010-2024 openpyxl

"""
Support for person/people list used by threaded comments.

The person list is stored at xl/persons/person.xml and contains
metadata about comment authors including GUIDs for identification.
"""

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import String
from openpyxl.descriptors.excel import Guid
from openpyxl.descriptors.sequence import Sequence

from openpyxl.xml.constants import XLPERSON_NS


class Person(Serialisable):
    """
    Represents a person/author in threaded comments.

    Attributes
    ----------
    displayName : str
        The display name of the person
    id : str
        A GUID uniquely identifying this person
    userId : str, optional
        Optional user identifier
    providerId : str, optional
        Optional provider identifier
    """

    tagname = "person"
    namespace = XLPERSON_NS

    displayName = String()
    id = Guid()
    userId = String(allow_none=True)
    providerId = String(allow_none=True)

    __attrs__ = ('displayName', 'id', 'userId', 'providerId')

    def __init__(self, displayName=None, id=None, userId=None, providerId=None):
        self.displayName = displayName
        self.id = id
        self.userId = userId
        self.providerId = providerId


class PersonList(Serialisable):
    """
    Container for all persons in a workbook.

    This is stored at xl/persons/person.xml in the XLSX archive.
    """

    tagname = "personList"
    namespace = XLPERSON_NS

    person = Sequence(expected_type=Person)

    _path = "/xl/persons/person.xml"
    mime_type = "application/vnd.ms-excel.person+xml"
    _rel_type = "http://schemas.microsoft.com/office/2017/10/relationships/person"

    __elements__ = ('person',)

    def __init__(self, person=()):
        self.person = person

    def __len__(self):
        return len(self.person)

    def __bool__(self):
        return bool(self.person)

    def __getitem__(self, key):
        """Get person by id (GUID)"""
        for p in self.person:
            if p.id == key:
                return p
        raise KeyError(f"Person with id {key} not found")

    def get(self, person_id, default=None):
        """Get person by id (GUID) with default"""
        try:
            return self[person_id]
        except KeyError:
            return default

    def append(self, person):
        """Add a person to the list"""
        self.person.append(person)

    @property
    def path(self):
        """Return path within the archive"""
        return self._path
