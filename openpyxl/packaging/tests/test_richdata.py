# Copyright (c) 2010-2024 openpyxl

import pytest
from io import BytesIO
from zipfile import ZipFile

from ..richdata import (
    RichDataManager,
    RICHVALUE_PATH,
    RICHVALUE_TYPES_PATH,
    RICHDATA_FILES,
)


class TestRichDataManager:

    def test_ctor(self):
        manager = RichDataManager()
        assert manager._files == {}
        assert not manager
        assert not manager.has_rich_data

    def test_read_empty(self):
        """Test reading from archive with no rich data"""
        buf = BytesIO()
        with ZipFile(buf, mode='w') as archive:
            archive.writestr('test.txt', b'test')

        buf.seek(0)
        with ZipFile(buf, mode='r') as archive:
            manager = RichDataManager()
            manager.read(archive, ['test.txt'])
            assert not manager

    def test_read_rich_data(self):
        """Test reading rich data files"""
        buf = BytesIO()
        with ZipFile(buf, mode='w') as archive:
            archive.writestr(RICHVALUE_PATH, b'<rdrichvalue/>')
            archive.writestr(RICHVALUE_TYPES_PATH, b'<types/>')

        buf.seek(0)
        with ZipFile(buf, mode='r') as archive:
            manager = RichDataManager()
            manager.read(archive, [RICHVALUE_PATH, RICHVALUE_TYPES_PATH])
            assert manager
            assert manager.has_rich_data
            assert RICHVALUE_PATH in manager._files
            assert manager._files[RICHVALUE_PATH] == b'<rdrichvalue/>'

    def test_write_rich_data(self):
        """Test writing preserved rich data"""
        from openpyxl.packaging.manifest import Manifest

        # Create manager with data
        manager = RichDataManager()
        manager._files[RICHVALUE_PATH] = b'<rdrichvalue/>'

        # Write to archive
        buf = BytesIO()
        manifest = Manifest()
        with ZipFile(buf, mode='w') as archive:
            manager.write(archive, manifest)

        # Verify
        buf.seek(0)
        with ZipFile(buf, mode='r') as archive:
            assert RICHVALUE_PATH in archive.namelist()
            assert archive.read(RICHVALUE_PATH) == b'<rdrichvalue/>'

    def test_round_trip(self):
        """Test round-trip preservation of rich data"""
        from openpyxl.packaging.manifest import Manifest

        # Create source archive
        buf1 = BytesIO()
        with ZipFile(buf1, mode='w') as archive:
            archive.writestr(RICHVALUE_PATH, b'<rdrichvalue>test data</rdrichvalue>')

        # Read from source
        buf1.seek(0)
        with ZipFile(buf1, mode='r') as archive:
            manager = RichDataManager()
            manager.read(archive, [RICHVALUE_PATH])

        # Write to destination
        buf2 = BytesIO()
        manifest = Manifest()
        with ZipFile(buf2, mode='w') as archive:
            manager.write(archive, manifest)

        # Verify content preserved exactly
        buf2.seek(0)
        with ZipFile(buf2, mode='r') as archive:
            content = archive.read(RICHVALUE_PATH)
            assert content == b'<rdrichvalue>test data</rdrichvalue>'

    def test_content_type(self):
        """Test correct content types are assigned"""
        from openpyxl.packaging.manifest import Manifest

        manager = RichDataManager()
        manager._files[RICHVALUE_PATH] = b'<test/>'

        buf = BytesIO()
        manifest = Manifest()
        with ZipFile(buf, mode='w') as archive:
            manager.write(archive, manifest)

        # Check manifest contains correct override
        assert any(
            o.PartName == "/" + RICHVALUE_PATH
            for o in manifest.Override
        )
