# Copyright (c) 2010-2024 openpyxl

"""
Support for rich data types (stocks, geography) in XLSX files.

Rich data types are stored in multiple files under xl/richData/:
- rdrichvalue.xml - Rich value data
- rdRichValueTypes.xml - Type definitions
- rdarray.xml - Array data
- rdRichValueStructure.xml - Structure definitions
- rdSupportingPropertyBag.xml - Supporting properties
- rdSupportingPropertyBagStructure.xml - Property bag structure
- rdRichValueWebImage.xml - Web images for rich data
- rdRichValueRel.xml - Relationships

This module provides binary blob preservation similar to VBA - the rich data
is read as raw bytes and written back unchanged to preserve round-trip fidelity.
"""


# Rich data file paths
RICHDATA_PATH = "xl/richData/"
RICHVALUE_PATH = RICHDATA_PATH + "rdrichvalue.xml"
RICHVALUE_TYPES_PATH = RICHDATA_PATH + "rdRichValueTypes.xml"
RICHVALUE_ARRAY_PATH = RICHDATA_PATH + "rdarray.xml"
RICHVALUE_STRUCTURE_PATH = RICHDATA_PATH + "rdRichValueStructure.xml"
RICHVALUE_PROPBAG_PATH = RICHDATA_PATH + "rdSupportingPropertyBag.xml"
RICHVALUE_PROPBAG_STRUCTURE_PATH = RICHDATA_PATH + "rdSupportingPropertyBagStructure.xml"
RICHVALUE_WEBIMAGE_PATH = RICHDATA_PATH + "rdRichValueWebImage.xml"
RICHVALUE_REL_PATH = RICHDATA_PATH + "rdRichValueRel.xml"

# Content types
RICHVALUE_TYPE = "application/vnd.ms-excel.rdrichvalue+xml"
RICHVALUE_TYPES_TYPE = "application/vnd.ms-excel.rdrichvaluetypes+xml"
RICHVALUE_ARRAY_TYPE = "application/vnd.ms-excel.rdarray+xml"
RICHVALUE_STRUCTURE_TYPE = "application/vnd.ms-excel.rdrichvaluestructure+xml"
RICHVALUE_PROPBAG_TYPE = "application/vnd.ms-excel.rdsupportingpropertybag+xml"
RICHVALUE_PROPBAG_STRUCTURE_TYPE = "application/vnd.ms-excel.rdsupportingpropertybagstructure+xml"
RICHVALUE_WEBIMAGE_TYPE = "application/vnd.ms-excel.rdRichValueWebImage+xml"
RICHVALUE_REL_TYPE = "application/vnd.ms-excel.rdrichvaluerel+xml"


# All known rich data file patterns
RICHDATA_FILES = [
    RICHVALUE_PATH,
    RICHVALUE_TYPES_PATH,
    RICHVALUE_ARRAY_PATH,
    RICHVALUE_STRUCTURE_PATH,
    RICHVALUE_PROPBAG_PATH,
    RICHVALUE_PROPBAG_STRUCTURE_PATH,
    RICHVALUE_WEBIMAGE_PATH,
    RICHVALUE_REL_PATH,
]


class RichDataManager:
    """
    Manager for rich data type preservation.

    This class preserves rich data files as binary blobs for round-trip fidelity.
    Rich data includes stocks, geography, and other linked data types introduced
    in Excel 2016+ and expanded in Office 365.

    Usage:
        # Reading
        manager = RichDataManager()
        manager.read(archive, valid_files)

        # Writing
        manager.write(archive, manifest)

    Note: This does not parse the rich data content - it only preserves it.
    Full parsing support for creating/modifying rich data would require
    substantial additional work.
    """

    def __init__(self):
        # Dict of path -> raw bytes
        self._files = {}

    def read(self, archive, valid_files):
        """
        Read all rich data files from the archive.

        Parameters
        ----------
        archive : ZipFile
            The XLSX archive
        valid_files : list
            List of valid file paths in the archive
        """
        for path in RICHDATA_FILES:
            if path in valid_files:
                try:
                    self._files[path] = archive.read(path)
                except KeyError:
                    pass  # File not in archive

    def write(self, archive, manifest):
        """
        Write all preserved rich data files to the archive.

        Parameters
        ----------
        archive : ZipFile
            The XLSX archive being written
        manifest : Manifest
            The content types manifest
        """
        from openpyxl.packaging.manifest import Override

        for path, content in self._files.items():
            archive.writestr(path, content)

            # Add content type based on path
            content_type = self._get_content_type(path)
            if content_type:
                # Create Override directly and add to manifest.Override list
                ct = Override(PartName="/" + path, ContentType=content_type)
                manifest.Override.append(ct)

    def _get_content_type(self, path):
        """Get content type for a rich data file path"""
        content_types = {
            RICHVALUE_PATH: RICHVALUE_TYPE,
            RICHVALUE_TYPES_PATH: RICHVALUE_TYPES_TYPE,
            RICHVALUE_ARRAY_PATH: RICHVALUE_ARRAY_TYPE,
            RICHVALUE_STRUCTURE_PATH: RICHVALUE_STRUCTURE_TYPE,
            RICHVALUE_PROPBAG_PATH: RICHVALUE_PROPBAG_TYPE,
            RICHVALUE_PROPBAG_STRUCTURE_PATH: RICHVALUE_PROPBAG_STRUCTURE_TYPE,
            RICHVALUE_WEBIMAGE_PATH: RICHVALUE_WEBIMAGE_TYPE,
            RICHVALUE_REL_PATH: RICHVALUE_REL_TYPE,
        }
        return content_types.get(path)

    def __bool__(self):
        """Return True if there is any rich data to preserve"""
        return bool(self._files)

    @property
    def has_rich_data(self):
        """Return True if there is any rich data to preserve"""
        return bool(self._files)
