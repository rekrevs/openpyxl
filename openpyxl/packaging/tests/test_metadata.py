# Copyright (c) 2010-2024 openpyxl

import pytest
from io import BytesIO

from openpyxl.packaging.metadata import (
    Metadata,
    MetadataTypes,
    MetadataType,
    MetadataRecord,
    MetadataBlock,
    CellMetadata,
    FutureMetadata,
    FutureMetadataBlock,
    read_metadata,
    METADATA_REL,
    METADATA_TYPE,
    XLDAPR_URI,
)
from openpyxl.xml.functions import fromstring, tostring


class TestMetadataType:

    def test_ctor(self):
        mt = MetadataType(name="XLDAPR", minSupportedVersion=120000, copy=True)
        assert mt.name == "XLDAPR"
        assert mt.minSupportedVersion == 120000
        assert mt.copy is True

    def test_from_tree(self):
        src = """
        <metadataType name="XLDAPR" minSupportedVersion="120000"
                      copy="1" pasteAll="1" pasteValues="1" merge="1"
                      splitFirst="1" rowColShift="1" clearFormats="1"
                      clearComments="1" assign="1" coerce="1"
                      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
        """
        node = fromstring(src)
        mt = MetadataType.from_tree(node)
        assert mt.name == "XLDAPR"
        assert mt.copy is True
        assert mt.coerce is True

    def test_to_tree(self):
        mt = MetadataType(name="XLDAPR", copy=True)
        tree = mt.to_tree()
        assert tree.get("name") == "XLDAPR"
        assert tree.get("copy") == "1"


class TestMetadataRecord:

    def test_ctor(self):
        rc = MetadataRecord(t=1, v=0)
        assert rc.t == 1
        assert rc.v == 0

    def test_from_tree(self):
        src = '<rc t="1" v="0" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'
        node = fromstring(src)
        rc = MetadataRecord.from_tree(node)
        assert rc.t == 1
        assert rc.v == 0


class TestCellMetadata:

    def test_from_tree(self):
        src = """
        <cellMetadata count="1" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <bk>
                <rc t="1" v="0"/>
            </bk>
        </cellMetadata>
        """
        node = fromstring(src)
        cm = CellMetadata.from_tree(node)
        assert cm.count == 1
        assert len(cm.bk) == 1
        assert cm.bk[0].rc[0].t == 1
        assert cm.bk[0].rc[0].v == 0


class TestFutureMetadata:

    def test_from_tree(self):
        src = """
        <futureMetadata name="XLDAPR" count="1"
                        xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <bk/>
        </futureMetadata>
        """
        node = fromstring(src)
        fm = FutureMetadata.from_tree(node)
        assert fm.name == "XLDAPR"
        assert fm.count == 1


class TestMetadata:

    def test_from_tree(self):
        src = """
        <metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <metadataTypes count="1">
                <metadataType name="XLDAPR" minSupportedVersion="120000"
                              copy="1" pasteAll="1" pasteValues="1"/>
            </metadataTypes>
            <futureMetadata name="XLDAPR" count="1">
                <bk/>
            </futureMetadata>
            <cellMetadata count="1">
                <bk>
                    <rc t="1" v="0"/>
                </bk>
            </cellMetadata>
        </metadata>
        """
        node = fromstring(src)
        metadata = Metadata.from_tree(node)
        assert metadata.metadataTypes is not None
        assert len(metadata.metadataTypes.metadataType) == 1
        assert metadata.metadataTypes.metadataType[0].name == "XLDAPR"
        assert len(metadata.futureMetadata) == 1
        assert metadata.cellMetadata is not None

    def test_to_tree(self):
        mt = MetadataType(name="XLDAPR", copy=True)
        mts = MetadataTypes(count=1, metadataType=[mt])
        fm = FutureMetadata(name="XLDAPR", count=1, bk=[FutureMetadataBlock()])
        rc = MetadataRecord(t=1, v=0)
        bk = MetadataBlock(rc=[rc])
        cm = CellMetadata(count=1, bk=[bk])

        metadata = Metadata(
            metadataTypes=mts,
            futureMetadata=[fm],
            cellMetadata=cm
        )

        tree = metadata.to_tree()
        # Tag may or may not have namespace prefix depending on context
        assert tree.tag.endswith("metadata")

    def test_path(self):
        metadata = Metadata()
        assert metadata.path == "/xl/metadata.xml"
        assert metadata.mime_type == METADATA_TYPE
        assert metadata.rel_type == METADATA_REL


class TestReadMetadata:

    def test_read_missing(self):
        """Reading from archive without metadata.xml returns None"""
        from zipfile import ZipFile
        buf = BytesIO()
        with ZipFile(buf, 'w') as zf:
            zf.writestr("test.xml", "<test/>")
        buf.seek(0)
        with ZipFile(buf, 'r') as archive:
            result = read_metadata(archive)
        assert result is None
