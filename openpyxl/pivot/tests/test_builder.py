# Copyright (c) 2010-2024 openpyxl

import pytest

from openpyxl import Workbook
from ..builder import PivotTableConfig, PivotTableBuilder


class TestPivotTableConfig:

    def test_ctor(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'Region'
        ws['B1'] = 'Sales'

        config = PivotTableConfig(
            source_worksheet=ws,
            source_range='A1:B10',
            destination_worksheet=ws,
            destination_cell='D1',
            name='TestPivot'
        )

        assert config.name == 'TestPivot'
        assert config.source_range == 'A1:B10'
        assert config.row_fields == []
        assert config.column_fields == []
        assert config.value_fields == []

    def test_add_row_field(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'Region'

        config = PivotTableConfig(
            source_worksheet=ws,
            source_range='A1:A10',
            destination_worksheet=ws,
            destination_cell='C1',
            name='Test'
        )
        config.add_row_field('Region')

        assert len(config.row_fields) == 1
        assert config.row_fields[0]['name'] == 'Region'

    def test_add_column_field(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'Product'

        config = PivotTableConfig(
            source_worksheet=ws,
            source_range='A1:A10',
            destination_worksheet=ws,
            destination_cell='C1',
            name='Test'
        )
        config.add_column_field('Product')

        assert len(config.column_fields) == 1
        assert config.column_fields[0]['name'] == 'Product'

    def test_add_value_field(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'Sales'

        config = PivotTableConfig(
            source_worksheet=ws,
            source_range='A1:A10',
            destination_worksheet=ws,
            destination_cell='C1',
            name='Test'
        )
        config.add_value_field('Sales', function='sum')

        assert len(config.value_fields) == 1
        assert config.value_fields[0]['name'] == 'Sales'
        assert config.value_fields[0]['function'] == 'sum'

    def test_get_source_headers(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'Region'
        ws['B1'] = 'Product'
        ws['C1'] = 'Sales'

        config = PivotTableConfig(
            source_worksheet=ws,
            source_range='A1:C10',
            destination_worksheet=ws,
            destination_cell='E1',
            name='Test'
        )
        headers = config.get_source_headers()

        assert headers == ['Region', 'Product', 'Sales']

    def test_get_field_index(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'Region'
        ws['B1'] = 'Product'
        ws['C1'] = 'Sales'

        config = PivotTableConfig(
            source_worksheet=ws,
            source_range='A1:C10',
            destination_worksheet=ws,
            destination_cell='E1',
            name='Test'
        )

        assert config.get_field_index('Region') == 0
        assert config.get_field_index('Product') == 1
        assert config.get_field_index('Sales') == 2

    def test_get_field_index_not_found(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'Region'

        config = PivotTableConfig(
            source_worksheet=ws,
            source_range='A1:A10',
            destination_worksheet=ws,
            destination_cell='C1',
            name='Test'
        )

        with pytest.raises(ValueError, match="Field 'NotFound' not found"):
            config.get_field_index('NotFound')

    def test_validate_success(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'Region'
        ws['B1'] = 'Sales'

        config = PivotTableConfig(
            source_worksheet=ws,
            source_range='A1:B10',
            destination_worksheet=ws,
            destination_cell='D1',
            name='TestPivot'
        )
        config.add_row_field('Region')
        config.add_value_field('Sales')

        assert config.validate() is True

    def test_validate_missing_name(self):
        wb = Workbook()
        ws = wb.active

        config = PivotTableConfig(
            source_worksheet=ws,
            source_range='A1:B10',
            destination_worksheet=ws,
            destination_cell='D1',
            name=''
        )

        with pytest.raises(ValueError, match="name is required"):
            config.validate()

    def test_validate_field_not_found(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'Region'

        config = PivotTableConfig(
            source_worksheet=ws,
            source_range='A1:A10',
            destination_worksheet=ws,
            destination_cell='C1',
            name='Test'
        )
        config.add_row_field('NotExists')

        with pytest.raises(ValueError, match="Field 'NotExists' not found"):
            config.validate()

    def test_method_chaining(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'Region'
        ws['B1'] = 'Product'
        ws['C1'] = 'Sales'

        config = PivotTableConfig(
            source_worksheet=ws,
            source_range='A1:C10',
            destination_worksheet=ws,
            destination_cell='E1',
            name='Test'
        )
        result = config.add_row_field('Region').add_column_field('Product').add_value_field('Sales')

        assert result is config
        assert len(config.row_fields) == 1
        assert len(config.column_fields) == 1
        assert len(config.value_fields) == 1


class TestPivotTableBuilder:

    def test_build_basic(self):
        wb = Workbook()
        ws = wb.active

        # Create test data
        ws['A1'] = 'Region'
        ws['B1'] = 'Product'
        ws['C1'] = 'Sales'
        ws['A2'] = 'East'
        ws['B2'] = 'Apple'
        ws['C2'] = 100

        builder = PivotTableBuilder(
            source_worksheet=ws,
            source_range='A1:C2',
            destination_worksheet=ws,
            destination_cell='E1',
            name='SalesPivot'
        )
        builder.add_row_field('Region')
        builder.add_value_field('Sales', function='sum')

        pivot = builder.build()

        assert pivot is not None
        assert pivot.name == 'SalesPivot'
        assert pivot.cache is not None

    def test_build_with_multiple_fields(self):
        wb = Workbook()
        ws = wb.active

        # Create test data
        ws['A1'] = 'Region'
        ws['B1'] = 'Product'
        ws['C1'] = 'Sales'
        ws['D1'] = 'Quantity'
        ws['A2'] = 'East'
        ws['B2'] = 'Apple'
        ws['C2'] = 100
        ws['D2'] = 10

        builder = PivotTableBuilder(
            source_worksheet=ws,
            source_range='A1:D2',
            destination_worksheet=ws,
            destination_cell='F1',
            name='DetailPivot'
        )
        builder.add_row_field('Region')
        builder.add_column_field('Product')
        builder.add_value_field('Sales', function='sum')
        builder.add_value_field('Quantity', function='sum')

        pivot = builder.build()

        assert pivot is not None
        assert pivot.name == 'DetailPivot'
        assert len(pivot.rowFields) == 1
        assert len(pivot.colFields) == 1
        assert len(pivot.dataFields) == 2

    def test_build_with_style(self):
        wb = Workbook()
        ws = wb.active

        ws['A1'] = 'Region'
        ws['B1'] = 'Sales'
        ws['A2'] = 'East'
        ws['B2'] = 100

        builder = PivotTableBuilder(
            source_worksheet=ws,
            source_range='A1:B2',
            destination_worksheet=ws,
            destination_cell='D1',
            name='StyledPivot'
        )
        builder.add_row_field('Region')
        builder.add_value_field('Sales')
        builder.set_style('PivotStyleMedium9')

        pivot = builder.build()

        assert pivot.pivotTableStyleInfo is not None
        assert pivot.pivotTableStyleInfo.name == 'PivotStyleMedium9'
