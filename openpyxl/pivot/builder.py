# Copyright (c) 2010-2024 openpyxl

"""
Builder classes for creating pivot tables programmatically.

This module provides a simpler API for creating pivot tables from scratch.
It handles the complex relationships between TableDefinition, CacheDefinition,
CacheRecords, and all the field configurations.

Usage:
    from openpyxl import Workbook
    from openpyxl.pivot.builder import PivotTableBuilder

    wb = Workbook()
    ws = wb.active

    # Add data
    ws.append(['Region', 'Product', 'Sales', 'Quantity'])
    ws.append(['East', 'Apple', 100, 10])
    ws.append(['West', 'Orange', 150, 15])
    # ...

    # Create pivot table
    builder = PivotTableBuilder(
        source_worksheet=ws,
        source_range='A1:D10',
        destination_worksheet=ws,
        destination_cell='F1',
        name='SalesPivot'
    )
    builder.add_row_field('Region')
    builder.add_row_field('Product')
    builder.add_value_field('Sales', function='sum')
    builder.add_value_field('Quantity', function='sum')

    pivot = builder.build()
    ws.add_pivot(pivot)
"""

from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter


class PivotTableConfig:
    """
    Configuration helper for pivot table creation.

    This stores the configuration for a pivot table before it's built.
    Since pivot table creation is complex, this class captures the user's
    intent which can then be used to build the actual pivot table structure.
    """

    def __init__(
        self,
        source_worksheet,
        source_range,
        destination_worksheet,
        destination_cell,
        name
    ):
        """
        Initialize pivot table configuration.

        Parameters
        ----------
        source_worksheet : Worksheet
            The worksheet containing the source data
        source_range : str
            The range containing source data (e.g., 'A1:D100')
        destination_worksheet : Worksheet
            The worksheet where the pivot table will be placed
        destination_cell : str
            The top-left cell of the pivot table (e.g., 'F1')
        name : str
            The name of the pivot table
        """
        self.source_worksheet = source_worksheet
        self.source_range = source_range
        self.destination_worksheet = destination_worksheet
        self.destination_cell = destination_cell
        self.name = name

        # Field configurations
        self.row_fields = []
        self.column_fields = []
        self.value_fields = []
        self.filter_fields = []

        # Style configuration
        self.style = None
        self.show_row_headers = True
        self.show_col_headers = True

    def add_row_field(self, field_name):
        """
        Add a field to the row area.

        Parameters
        ----------
        field_name : str
            The name of the field (column header in source data)
        """
        self.row_fields.append({'name': field_name})
        return self

    def add_column_field(self, field_name):
        """
        Add a field to the column area.

        Parameters
        ----------
        field_name : str
            The name of the field (column header in source data)
        """
        self.column_fields.append({'name': field_name})
        return self

    def add_value_field(self, field_name, function='sum', number_format=None):
        """
        Add a field to the values area.

        Parameters
        ----------
        field_name : str
            The name of the field (column header in source data)
        function : str
            The aggregation function: 'sum', 'count', 'average', 'max',
            'min', 'product', 'countNums', 'stdDev', 'stdDevP', 'var', 'varP'
        number_format : str, optional
            Number format for the values
        """
        self.value_fields.append({
            'name': field_name,
            'function': function,
            'number_format': number_format
        })
        return self

    def add_filter_field(self, field_name):
        """
        Add a field to the filter (page) area.

        Parameters
        ----------
        field_name : str
            The name of the field (column header in source data)
        """
        self.filter_fields.append({'name': field_name})
        return self

    def set_style(self, style_name):
        """
        Set the pivot table style.

        Parameters
        ----------
        style_name : str
            The name of the pivot table style (e.g., 'PivotStyleMedium9')
        """
        self.style = style_name
        return self

    def get_source_headers(self):
        """
        Get the headers from the source range.

        Returns
        -------
        list
            List of header names from the first row of the source range
        """
        from openpyxl.utils import range_boundaries

        min_col, min_row, max_col, max_row = range_boundaries(self.source_range)
        headers = []
        for col in range(min_col, max_col + 1):
            cell = self.source_worksheet.cell(row=min_row, column=col)
            headers.append(cell.value)
        return headers

    def get_field_index(self, field_name):
        """
        Get the index of a field by name.

        Parameters
        ----------
        field_name : str
            The name of the field

        Returns
        -------
        int
            The 0-based index of the field

        Raises
        ------
        ValueError
            If the field name is not found
        """
        headers = self.get_source_headers()
        try:
            return headers.index(field_name)
        except ValueError:
            raise ValueError(f"Field '{field_name}' not found in source headers: {headers}")

    def validate(self):
        """
        Validate the configuration.

        Raises
        ------
        ValueError
            If the configuration is invalid
        """
        if not self.name:
            raise ValueError("Pivot table name is required")

        if not self.source_range:
            raise ValueError("Source range is required")

        # Validate all field names exist
        headers = self.get_source_headers()

        for field in self.row_fields + self.column_fields + self.filter_fields:
            if field['name'] not in headers:
                raise ValueError(f"Field '{field['name']}' not found in source headers: {headers}")

        for field in self.value_fields:
            if field['name'] not in headers:
                raise ValueError(f"Field '{field['name']}' not found in source headers: {headers}")

        return True


class PivotTableBuilder(PivotTableConfig):
    """
    Builder for creating pivot tables.

    This class extends PivotTableConfig with the ability to build
    the actual TableDefinition and CacheDefinition objects.

    Note: Full pivot table creation is complex and requires creating
    multiple interconnected objects. This builder provides a starting
    point but may not cover all advanced scenarios.
    """

    def build(self):
        """
        Build and return the pivot table definition.

        This creates a TableDefinition and associated CacheDefinition
        based on the configuration.

        Returns
        -------
        TableDefinition
            The pivot table definition object

        Note
        ----
        After calling build(), you should add the pivot table to the
        destination worksheet using ws.add_pivot(pivot).

        The pivot table will need an associated cache which is managed
        by the workbook when saving.
        """
        self.validate()

        from .table import (
            TableDefinition,
            Location,
            PivotField,
            RowColField,
            DataField,
            PageField,
            PivotTableStyle,
        )
        from .cache import CacheDefinition, CacheSource, WorksheetSource, CacheField

        # Get source information
        headers = self.get_source_headers()
        from openpyxl.utils import range_boundaries
        min_col, min_row, max_col, max_row = range_boundaries(self.source_range)
        num_fields = len(headers)

        # Create cache source
        ws_source = WorksheetSource(
            ref=self.source_range,
            sheet=self.source_worksheet.title
        )
        cache_source = CacheSource(type='worksheet', worksheetSource=ws_source)

        # Create cache fields for each column
        cache_fields = []
        for header in headers:
            cache_fields.append(CacheField(name=header, numFmtId=0))

        # Create cache definition
        cache = CacheDefinition(
            cacheSource=cache_source,
            cacheFields=cache_fields
        )

        # Create pivot fields for the table definition
        pivot_fields = []
        for i, header in enumerate(headers):
            # Check if this field is used in row/column/filter
            is_row = any(f['name'] == header for f in self.row_fields)
            is_col = any(f['name'] == header for f in self.column_fields)
            is_filter = any(f['name'] == header for f in self.filter_fields)
            is_value = any(f['name'] == header for f in self.value_fields)

            axis = None
            if is_row:
                axis = 'axisRow'
            elif is_col:
                axis = 'axisCol'
            elif is_filter:
                axis = 'axisPage'

            # dataField=True if this field is used as a value
            pivot_fields.append(PivotField(
                axis=axis,
                dataField=is_value,
                showAll=False
            ))

        # Create row fields
        row_fields = []
        for field in self.row_fields:
            idx = self.get_field_index(field['name'])
            row_fields.append(RowColField(x=idx))

        # Create column fields
        col_fields = []
        for field in self.column_fields:
            idx = self.get_field_index(field['name'])
            col_fields.append(RowColField(x=idx))

        # Create data (value) fields
        data_fields = []
        for field in self.value_fields:
            idx = self.get_field_index(field['name'])
            func = field.get('function', 'sum')
            data_fields.append(DataField(
                name=f"Sum of {field['name']}" if func == 'sum' else f"{func.title()} of {field['name']}",
                fld=idx,
                subtotal=func
            ))

        # Create page (filter) fields
        page_fields = []
        for field in self.filter_fields:
            idx = self.get_field_index(field['name'])
            page_fields.append(PageField(fld=idx))

        # Calculate destination location
        dest_row, dest_col = coordinate_to_tuple(self.destination_cell)
        # Default location size (will be recalculated by Excel)
        location_ref = f"{self.destination_cell}:{get_column_letter(dest_col + 3)}{dest_row + 10}"
        location = Location(
            ref=location_ref,
            firstHeaderRow=1,
            firstDataRow=2,
            firstDataCol=1
        )

        # Create style if specified
        style_info = None
        if self.style:
            style_info = PivotTableStyle(name=self.style, showRowHeaders=True, showColHeaders=True)

        # Create table definition
        # Note: Empty sequences should be () not None for NestedSequence fields
        pivot = TableDefinition(
            name=self.name,
            cacheId=0,  # Will be set by workbook when saving
            dataCaption="Values",
            location=location,
            pivotFields=pivot_fields,
            rowFields=row_fields if row_fields else (),
            colFields=col_fields if col_fields else (),
            dataFields=data_fields if data_fields else (),
            pageFields=page_fields if page_fields else (),
            pivotTableStyleInfo=style_info
        )

        # Associate the cache
        pivot.cache = cache

        return pivot
