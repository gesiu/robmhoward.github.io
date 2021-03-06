declare module Excel {
    /**
     *
     * Represents the Excel application that manages the workbook.
     */
    class Application extends OfficeExtension.ClientObject {
        private m_calculationMode;
        /**
         *
         * Returns the calculation mode used in the workbook. See Excel.CalculationMode for details. Read-only.
         */
        calculationMode: string;
        /**
         *
         * Recalculate all currently opened workbooks in Excel.
         *
         * @param calculationType Specifies the calculation type to use. See Excel.CalculationType for details.
         */
        calculate(calculationType: string): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.Application;
    }
    /**
     *
     * Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc.
     */
    class Workbook extends OfficeExtension.ClientObject {
        private m_application;
        private m_bindings;
        private m_names;
        private m_tables;
        private m_worksheets;
        /**
         *
         * Represents Excel application instance that contains this workbook. Read-only.
         */
        application: Excel.Application;
        /**
         *
         * Represents a collection of bindings that are part of the workbook. Read-only.
         */
        bindings: Excel.BindingCollection;
        /**
         *
         * Represents a collection of workbook scoped named items (named ranges and constants). Read-only.
         */
        names: Excel.NamedItemCollection;
        /**
         *
         * Represents a collection of tables associated with the workbook. Read-only.
         */
        tables: Excel.TableCollection;
        /**
         *
         * Represents a collection of worksheets associated with the workbook. Read-only.
         */
        worksheets: Excel.WorksheetCollection;
        /**
         *
         * Gets the currently selected range from the workbook.
         *
         */
        getSelectedRange(): Excel.Range;
        _GetObjectByReferenceId(bstrReferenceId: string): OfficeExtension.ClientResult<any>;
        _GetObjectTypeNameByReferenceId(bstrReferenceId: string): OfficeExtension.ClientResult<string>;
        _GetReferenceCount(): OfficeExtension.ClientResult<number>;
        _RemoveAllReferences(): void;
        _RemoveReference(bstrReferenceId: string): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.Workbook;
    }
    /**
     *
     * An Excel worksheet is a grid of cells. It can contain data, tables, charts, etc.
     */
    class Worksheet extends OfficeExtension.ClientObject {
        private m_charts;
        private m_id;
        private m_name;
        private m_position;
        private m_tables;
        private m_visibility;
        /**
         *
         * Returns collection of charts that are part of the worksheet. Read-only.
         */
        charts: Excel.ChartCollection;
        /**
         *
         * Collection of tables that are part of the worksheet. Read-only.
         */
        tables: Excel.TableCollection;
        /**
         *
         * Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved. Read-only.
         */
        id: string;
        /**
         *
         * The display name of the worksheet.
         */
        name: string;
        /**
         *
         * The zero-based position of the worksheet within the workbook.
         */
        position: number;
        /**
         *
         * The Visibility of the worksheet, Read-only.
         */
        visibility: string;
        /**
         *
         * Activate the worksheet in the Excel UI.
         *
         */
        activate(): void;
        /**
         *
         * Deletes the worksheet from the workbook.
         *
         */
        delete(): void;
        /**
         *
         * Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid.
         *
         * @param row The row number of the cell to be retrieved. Zero-indexed.
         * @param column the column number of the cell to be retrieved. Zero-indexed.
         */
        getCell(row: number, column: number): Excel.Range;
        /**
         *
         * Gets the range object specified by the address or name.
         *
         * @param address The address or the name of the range. If not specified, the entire worksheet range is returned.
         */
        getRange(address?: string): Excel.Range;
        /**
         *
         * The used range is the smallest range than encompasses any cells that have a value or formatting assigned to them. If the worksheet is blank, this function will return the top left cell.
         *
         */
        getUsedRange(): Excel.Range;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.Worksheet;
    }
    /**
     *
     * Represents a collection of worksheet objects that are part of the workbook.
     */
    class WorksheetCollection extends OfficeExtension.ClientObject {
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.Worksheet>;
        /**
         *
         * Adds a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets. If you wish to activate the newly added worksheet, call ".activate() on it.
         *
         * @param name The name of the worksheet to be added. If specified, name should be unqiue. If not specified, Excel determines the name of the new worksheet.
         */
        add(name?: string): Excel.Worksheet;
        /**
         *
         * Gets the currently active worksheet in the workbook.
         *
         */
        getActiveWorksheet(): Excel.Worksheet;
        /**
         *
         * Gets a worksheet object using its Name or ID.
         *
         * @param index The Name or ID of the worksheet.
         */
        getItem(key: string): Excel.Worksheet;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.WorksheetCollection;
    }
    /**
     *
     * Range represents a set of one or more contiguous cells such as a cell, a row, a column, block of cells, etc.
     */
    class Range extends OfficeExtension.ClientObject {
        private m_address;
        private m_addressLocal;
        private m_cellCount;
        private m_columnCount;
        private m_columnIndex;
        private m_format;
        private m_formulas;
        private m_formulasLocal;
        private m_numberFormat;
        private m_rowCount;
        private m_rowIndex;
        private m_text;
        private m_valueTypes;
        private m_values;
        private m_worksheet;
        private m__ReferenceId;
        /**
         *
         * Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties. Read-only.
         */
        format: Excel.RangeFormat;
        /**
         *
         * The worksheet containing the current range. Read-only.
         */
        worksheet: Excel.Worksheet;
        /**
         *
         * Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. Sheet1!A1:B4). Read-only.
         */
        address: string;
        /**
         *
         * Represents range reference for the specified range in the language of the user. Read-only.
         */
        addressLocal: string;
        /**
         *
         * Number of cells in the range. Read-only.
         */
        cellCount: number;
        /**
         *
         * Represents the total number of columns in the range. Read-only.
         */
        columnCount: number;
        /**
         *
         * Represents the column number of the first cell in the range. Zero-indexed. Read-only.
         */
        columnIndex: number;
        /**
         *
         * Represents the formula in A1-style notation.
         */
        formulas: Array<Array<any>>;
        /**
         *
         * Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.
         */
        formulasLocal: Array<Array<any>>;
        /**
         *
         * Represents Excel's number format code for the given cell.
         */
        numberFormat: Array<Array<any>>;
        /**
         *
         * Returns the total number of rows in the range. Read-only.
         */
        rowCount: number;
        /**
         *
         * Returns the row number of the first cell in the range. Zero-indexed. Read-only.
         */
        rowIndex: number;
        /**
         *
         * Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.
         */
        text: Array<Array<any>>;
        /**
         *
         * Represents the type of data of each cell. Read-only.
         */
        valueTypes: Array<Array<string>>;
        /**
         *
         * Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
         */
        values: Array<Array<any>>;
        _ReferenceId: string;
        /**
         *
         * Clear range values, format, fill, border, etc.
         *
         * @param applyTo Determines the type of clear action. See Excel.ClearApplyTo for details.
         */
        clear(applyTo?: string): void;
        /**
         *
         * Deletes the cells associated with the range.
         *
         * @param shift Specifies which way to shift the cells. See Excel.DeleteShiftDirection for details.
         */
        delete(shift: string): void;
        /**
         *
         * Gets the smallest range object that encompasses the given ranges. For example, the GetBoundingRect of "B2:C5" and "D10:E15" is "B2:E16".
         *
         * @param anotherRange The range object or address or range name.
         */
        getBoundingRect(anotherRange: Excel.Range | string): Excel.Range;
        /**
         *
         * Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it's stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.
         *
         * @param row Row number of the cell to be retrieved. Zero-indexed.
         * @param column Column number of the cell to be retrieved. Zero-indexed.
         */
        getCell(row: number, column: number): Excel.Range;
        /**
         *
         * Gets a column contained in the range.
         *
         * @param column Column number of the range to be retrieved. Zero-indexed.
         */
        getColumn(column: number): Excel.Range;
        /**
         *
         * Gets an object that represents the entire column of the range.
         *
         */
        getEntireColumn(): Excel.Range;
        /**
         *
         * Gets an object that represents the entire row of the range.
         *
         */
        getEntireRow(): Excel.Range;
        /**
         *
         * Gets the range object that represents the rectangular intersection of the given ranges.
         *
         * @param anotherRange The range object or range address that will be used to determine the intersection of ranges.
         */
        getIntersection(anotherRange: Excel.Range | string): Excel.Range;
        /**
         *
         * Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".
         *
         */
        getLastCell(): Excel.Range;
        /**
         *
         * Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".
         *
         */
        getLastColumn(): Excel.Range;
        /**
         *
         * Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".
         *
         */
        getLastRow(): Excel.Range;
        /**
         *
         * Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an exception will be thrown.
         *
         * @param rowOffset The number of rows (positive, negative, or 0) by which the range is to be offset. Positive values are offset downward, and negative values are offset upward.
         * @param columnOffset The number of columns (positive, negative, or 0) by which the range is to be offset. Positive values are offset to the right, and negative values are offset to the left.
         */
        getOffsetRange(rowOffset: number, columnOffset: number): Excel.Range;
        /**
         *
         * Gets a row contained in the range.
         *
         * @param row Row number of the range to be retrieved. Zero-indexed.
         */
        getRow(row: number): Excel.Range;
        /**
         *
         * Returns the used range of the given range object.
         *
         */
        getUsedRange(): Excel.Range;
        /**
         *
         * Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new Range object at the now blank space.
         *
         * @param shift Specifies which way to shift the cells. See Excel.InsertShiftDirection for details.
         */
        insert(shift: string): Excel.Range;
        /**
         *
         * Selects the specified range in the Excel UI.
         *
         */
        select(): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.Range;
        _initReferenceId(value: string): void;
    }
    /**
     *
     * A collection of all the nameditem objects that are part of the workbook.
     */
    class NamedItemCollection extends OfficeExtension.ClientObject {
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.NamedItem>;
        /**
         *
         * Gets a nameditem object using its name
         *
         * @param name nameditem name.
         */
        getItem(name: string): Excel.NamedItem;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.NamedItemCollection;
    }
    /**
     *
     * Represents a defined name for a range of cells or value. Names can be primitive named objects (as seen in the type below), range object, reference to a range. This object can be used to obtain range object associated with names.
     */
    class NamedItem extends OfficeExtension.ClientObject {
        private m_name;
        private m_type;
        private m_value;
        private m_visible;
        private m__Id;
        /**
         *
         * The name of the object. Read-only.
         */
        name: string;
        /**
         *
         * Indicates what type of reference is associated with the name. See Excel.NamedItemType for details. Read-only.
         */
        type: string;
        /**
         *
         * Represents the formula that the name is defined to refer to. E.g. =Sheet14!$B$2:$H$12, =4.75, etc. Read-only.
         */
        value: any;
        /**
         *
         * Specifies whether the object is visible or not.
         */
        visible: boolean;
        _Id: string;
        /**
         *
         * Returns the range object that is associated with the name. Throws an exception if the named item's type is not a range.
         *
         */
        getRange(): Excel.Range;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.NamedItem;
    }
    /**
     *
     * Represents an Office.js binding that is defined in the workbook.
     */
    class Binding extends OfficeExtension.ClientObject {
        private m_id;
        private m_type;
        /**
         *
         * Represents binding identifier. Read-only.
         */
        id: string;
        /**
         *
         * Returns the type of the binding. See Excel.BindingType for details. Read-only.
         */
        type: string;
        /**
         *
         * Returns the range represented by the binding. Will throw an error if binding is not of the correct type.
         *
         */
        getRange(): Excel.Range;
        /**
         *
         * Returns the table represented by the binding. Will throw an error if binding is not of the correct type.
         *
         */
        getTable(): Excel.Table;
        /**
         *
         * Returns the text represented by the binding. Will throw an error if binding is not of the correct type.
         *
         */
        getText(): OfficeExtension.ClientResult<string>;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.Binding;
    }
    /**
     *
     * Represents the collection of all the binding objects that are part of the workbook.
     */
    class BindingCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.Binding>;
        /**
         *
         * Returns the number of bindings in the collection. Read-only.
         */
        count: number;
        /**
         *
         * Gets a binding object by ID.
         *
         * @param id Id of the binding object to be retrieved.
         */
        getItem(id: string): Excel.Binding;
        /**
         *
         * Gets a binding object based on its position in the items array.
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Excel.Binding;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.BindingCollection;
    }
    /**
     *
     * Represents a collection of all the tables that are part of the workbook.
     */
    class TableCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.Table>;
        /**
         *
         * Returns the number of tables in the workbook. Read-only.
         */
        count: number;
        /**
         *
         * Create a new table. The range source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.
         *
         * @param address Address or name of the range object representing the data source. If the address does not contain a sheet name, the currently-active sheet is used.
         * @param hasHeaders Boolean value that indicates whether the data being imported has column labels. If the source does not contain headers (i.e,. when this property set to false), Excel will automatically generate header shifting the data down by one row.
         */
        add(address: string, hasHeaders: boolean): Excel.Table;
        /**
         *
         * Gets a table by Name or ID.
         *
         * @param id Name or ID of the table to be retrieved.
         */
        getItem(key: number | string): Excel.Table;
        /**
         *
         * Gets a table based on its position in the collection.
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Excel.Table;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.TableCollection;
    }
    /**
     *
     * Represents an Excel table.
     */
    class Table extends OfficeExtension.ClientObject {
        private m_columns;
        private m_id;
        private m_name;
        private m_rows;
        private m_showHeaders;
        private m_showTotals;
        private m_style;
        /**
         *
         * Represents a collection of all the columns in the table. Read-only.
         */
        columns: Excel.TableColumnCollection;
        /**
         *
         * Represents a collection of all the rows in the table. Read-only.
         */
        rows: Excel.TableRowCollection;
        /**
         *
         * Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.
         */
        id: number;
        /**
         *
         * Name of the table.
         */
        name: string;
        /**
         *
         * Indicates whether the header row is visible or not. This value can be set to show or remove the header row.
         */
        showHeaders: boolean;
        /**
         *
         * Indicates whether the total row is visible or not. This value can be set to show or remove the total row.
         */
        showTotals: boolean;
        /**
         *
         * Constant value that represents the Table style. Possible values are: TableStyleLight1 thru TableStyleLight21, TableStyleMedium1 thru TableStyleMedium28, TableStyleStyleDark1 thru TableStyleStyleDark11. A custom user-defined style present in the workbook can also be specified.
         */
        style: string;
        /**
         *
         * Deletes the table.
         *
         */
        delete(): void;
        /**
         *
         * Gets the range object associated with the data body of the table.
         *
         */
        getDataBodyRange(): Excel.Range;
        /**
         *
         * Gets the range object associated with header row of the table.
         *
         */
        getHeaderRowRange(): Excel.Range;
        /**
         *
         * Gets the range object associated with the entire table.
         *
         */
        getRange(): Excel.Range;
        /**
         *
         * Gets the range object associated with totals row of the table.
         *
         */
        getTotalRowRange(): Excel.Range;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.Table;
    }
    /**
     *
     * Represents a collection of all the columns that are part of the table.
     */
    class TableColumnCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.TableColumn>;
        /**
         *
         * Returns the number of columns in the table. Read-only.
         */
        count: number;
        /**
         *
         * Adds a new column to the table.
         *
         * @param index Specifies the relative position of the new column. The previous column at this position is shifted to the right. The index value should be equal to or less than the last column's index value, so it cannot be used to append a column at the end of the table. Zero-indexed.
         * @param values A 2-dimensional array of unformatted values of the table column.
         */
        add(index: number, values?: Array<Array<boolean | string | number>> | boolean | string | number): Excel.TableColumn;
        /**
         *
         * Gets a column object by Name or ID.
         *
         * @param id Column Name or ID.
         */
        getItem(key: number | string): Excel.TableColumn;
        /**
         *
         * Gets a column based on its position in the collection.
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Excel.TableColumn;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.TableColumnCollection;
    }
    /**
     *
     * Represents a column in a table.
     */
    class TableColumn extends OfficeExtension.ClientObject {
        private m_id;
        private m_index;
        private m_name;
        private m_values;
        /**
         *
         * Returns a unique key that identifies the column within the table. Read-only.
         */
        id: number;
        /**
         *
         * Returns the index number of the column within the columns collection of the table. Zero-indexed. Read-only.
         */
        index: number;
        /**
         *
         * Returns the name of the table column. Read-only.
         */
        name: string;
        /**
         *
         * Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
         */
        values: Array<Array<any>>;
        /**
         *
         * Deletes the column from the table.
         *
         */
        delete(): void;
        /**
         *
         * Gets the range object associated with the data body of the column.
         *
         */
        getDataBodyRange(): Excel.Range;
        /**
         *
         * Gets the range object associated with the header row of the column.
         *
         */
        getHeaderRowRange(): Excel.Range;
        /**
         *
         * Gets the range object associated with the entire column.
         *
         */
        getRange(): Excel.Range;
        /**
         *
         * Gets the range object associated with the totals row of the column.
         *
         */
        getTotalRowRange(): Excel.Range;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.TableColumn;
    }
    /**
     *
     * Represents a collection of all the rows that are part of the table.
     */
    class TableRowCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.TableRow>;
        /**
         *
         * Returns the number of rows in the table. Read-only.
         */
        count: number;
        /**
         *
         * Adds a new row to the table.
         *
         * @param index Specifies the relative position of the new row. If null, the addition happens at the end. Any rows below the inserted row are shifted downwards. Zero-indexed.
         * @param values A 2-dimensional array of unformatted values of the table row.
         */
        add(index?: number, values?: Array<Array<boolean | string | number>> | boolean | string | number): Excel.TableRow;
        /**
         *
         * Gets a row based on its position in the collection.
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Excel.TableRow;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.TableRowCollection;
    }
    /**
     *
     * Represents a row in a table.
     */
    class TableRow extends OfficeExtension.ClientObject {
        private m_index;
        private m_values;
        /**
         *
         * Returns the index number of the row within the rows collection of the table. Zero-indexed. Read-only.
         */
        index: number;
        /**
         *
         * Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.
         */
        values: Array<Array<any>>;
        /**
         *
         * Deletes the row from the table.
         *
         */
        delete(): void;
        /**
         *
         * Returns the range object associated with the entire row.
         *
         */
        getRange(): Excel.Range;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.TableRow;
    }
    /**
     *
     * A format object encapsulating the range's font, fill, borders, alignment, and other properties.
     */
    class RangeFormat extends OfficeExtension.ClientObject {
        private m_borders;
        private m_fill;
        private m_font;
        private m_horizontalAlignment;
        private m_verticalAlignment;
        private m_wrapText;
        /**
         *
         * Collection of border objects that apply to the overall range selected Read-only.
         */
        borders: Excel.RangeBorderCollection;
        /**
         *
         * Returns the fill object defined on the overall range. Read-only.
         */
        fill: Excel.RangeFill;
        /**
         *
         * Returns the font object defined on the overall range selected Read-only.
         */
        font: Excel.RangeFont;
        /**
         *
         * Represents the horizontal alignment for the specified object. See Excel.HorizontalAlignment for details.
         */
        horizontalAlignment: string;
        /**
         *
         * Represents the vertical alignment for the specified object. See Excel.VerticalAlignment for details.
         */
        verticalAlignment: string;
        /**
         *
         * Indicates if Excel wraps the text in the object. A null value indicates that the entire range doesn't have uniform wrap setting
         */
        wrapText: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.RangeFormat;
    }
    /**
     *
     * Represents the background of a range object.
     */
    class RangeFill extends OfficeExtension.ClientObject {
        private m_color;
        /**
         *
         * HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")
         */
        color: string;
        /**
         *
         * Resets the range background.
         *
         */
        clear(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.RangeFill;
    }
    /**
     *
     * Represents the border of an object.
     */
    class RangeBorder extends OfficeExtension.ClientObject {
        private m_color;
        private m_sideIndex;
        private m_style;
        private m_weight;
        /**
         *
         * HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
         */
        color: string;
        /**
         *
         * Constant value that indicates the specific side of the border. See Excel.BorderIndex for details. Read-only.
         */
        sideIndex: string;
        /**
         *
         * One of the constants of line style specifying the line style for the border. See Excel.BorderLineStyle for details.
         */
        style: string;
        /**
         *
         * Specifies the weight of the border around a range. See Excel.BorderWeight for details.
         */
        weight: string;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.RangeBorder;
    }
    /**
     *
     * Represents the border objects that make up range border.
     */
    class RangeBorderCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.RangeBorder>;
        /**
         *
         * Number of border objects in the collection. Read-only.
         */
        count: number;
        /**
         *
         * Gets a border object using its name
         *
         * @param index Index value of the border object to be retrieved. See Excel.BorderIndex for details.
         */
        getItem(index: string): Excel.RangeBorder;
        /**
         *
         * Gets a border object using its index
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Excel.RangeBorder;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.RangeBorderCollection;
    }
    /**
     *
     * This object represents the font attributes (font name, font size, color, etc.) for an object.
     */
    class RangeFont extends OfficeExtension.ClientObject {
        private m_bold;
        private m_color;
        private m_italic;
        private m_name;
        private m_size;
        private m_underline;
        /**
         *
         * Represents the bold status of font.
         */
        bold: boolean;
        /**
         *
         * HTML color code representation of the text color. E.g. #FF0000 represents Red.
         */
        color: string;
        /**
         *
         * Represents the italic status of the font.
         */
        italic: boolean;
        /**
         *
         * Font name (e.g. "Calibri")
         */
        name: string;
        /**
         *
         * Font size.
         */
        size: number;
        /**
         *
         * Type of underline applied to the font. See Excel.RangeUnderlineStyle for details.
         */
        underline: string;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.RangeFont;
    }
    /**
     *
     * A collection of all the chart objects on a worksheet.
     */
    class ChartCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.Chart>;
        /**
         *
         * Returns the number of charts in the worksheet. Read-only.
         */
        count: number;
        /**
         *
         * Creates a new chart.
         *
         * @param type Represents the type of a chart. See Excel.ChartType for details.
         * @param sourceData The Range object corresponding to the source data.
         * @param seriesBy Specifies the way columns or rows are used as data series on the chart. See Excel.ChartSeriesBy for details.
         */
        add(type: string, sourceData: Excel.Range, seriesBy?: string): Excel.Chart;
        /**
         *
         * Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.
         *
         * @param name Name of the chart to be retrieved.
         */
        getItem(name: string): Excel.Chart;
        /**
         *
         * Gets a chart based on its position in the collection.
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Excel.Chart;
        _GetItem(id: string): Excel.Chart;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartCollection;
    }
    /**
     *
     * Represents a chart object in a workbook.
     */
    class Chart extends OfficeExtension.ClientObject {
        private m_axes;
        private m_dataLabels;
        private m_format;
        private m_height;
        private m_left;
        private m_legend;
        private m_name;
        private m_series;
        private m_title;
        private m_top;
        private m_width;
        private m__Id;
        /**
         *
         * Represents chart axes. Read-only.
         */
        axes: Excel.ChartAxes;
        /**
         *
         * Represents the datalabels on the chart. Read-only.
         */
        dataLabels: Excel.ChartDataLabels;
        /**
         *
         * Encapsulates the format properties for the chart area. Read-only.
         */
        format: Excel.ChartAreaFormat;
        /**
         *
         * Represents the legend for the chart. Read-only.
         */
        legend: Excel.ChartLegend;
        /**
         *
         * Represents either a single series or collection of series in the chart. Read-only.
         */
        series: Excel.ChartSeriesCollection;
        /**
         *
         * Represents the title of the specified chart, including the text, visibility, position and formating of the title. Read-only.
         */
        title: Excel.ChartTitle;
        /**
         *
         * Represents the height, in points, of the chart object.
         */
        height: number;
        /**
         *
         * The distance, in points, from the left side of the chart to the worksheet origin.
         */
        left: number;
        /**
         *
         * Represents the name of a chart object.
         */
        name: string;
        /**
         *
         * Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
         */
        top: number;
        /**
         *
         * Represents the width, in points, of the chart object.
         */
        width: number;
        _Id: string;
        /**
         *
         * Deletes the chart object.
         *
         */
        delete(): void;
        /**
         *
         * Resets the source data for the chart.
         *
         * @param sourceData The Range object corresponding to the source data.
         * @param seriesBy Specifies the way columns or rows are used as data series on the chart. Can be one of the following: Auto (default), Rows, Columns. See Excel.ChartSeriesBy for details.
         */
        setData(sourceData: Excel.Range, seriesBy?: string): void;
        /**
         *
         * Positions the chart relative to cells on the worksheet.
         *
         * @param startCell The start cell. This is where the chart will be moved to. The start cell is the top-left or top-right cell, depending on the user's right-to-left display settings.
         * @param endCell (Optional) The end cell. If specified, the chart's width and height will be set to fully cover up this cell/range.
         */
        setPosition(startCell: Excel.Range | string, endCell?: Excel.Range | string): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.Chart;
    }
    /**
     *
     * Encapsulates the format properties for the overall chart area.
     */
    class ChartAreaFormat extends OfficeExtension.ClientObject {
        private m_fill;
        private m_font;
        /**
         *
         * Represents the fill format of an object, which includes background formatting information. Read-only.
         */
        fill: Excel.ChartFill;
        /**
         *
         * Represents the font attributes (font name, font size, color, etc.) for the current object. Read-only.
         */
        font: Excel.ChartFont;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartAreaFormat;
    }
    /**
     *
     * Represents a collection of chart series.
     */
    class ChartSeriesCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.ChartSeries>;
        /**
         *
         * Returns the number of series in the collection. Read-only.
         */
        count: number;
        /**
         *
         * Retrieves a series based on its position in the collection
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Excel.ChartSeries;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartSeriesCollection;
    }
    /**
     *
     * Represents a series in a chart.
     */
    class ChartSeries extends OfficeExtension.ClientObject {
        private m_format;
        private m_name;
        private m_points;
        /**
         *
         * Represents the formatting of a chart series, which includes fill and line formatting. Read-only.
         */
        format: Excel.ChartSeriesFormat;
        /**
         *
         * Represents a collection of all points in the series. Read-only.
         */
        points: Excel.ChartPointsCollection;
        /**
         *
         * Represents the name of a series in a chart.
         */
        name: string;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartSeries;
    }
    /**
     *
     * encapsulates the format properties for the chart series
     */
    class ChartSeriesFormat extends OfficeExtension.ClientObject {
        private m_fill;
        private m_line;
        /**
         *
         * Represents the fill format of a chart series, which includes background formating information. Read-only.
         */
        fill: Excel.ChartFill;
        /**
         *
         * Represents line formatting. Read-only.
         */
        line: Excel.ChartLineFormat;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartSeriesFormat;
    }
    /**
     *
     * A collection of all the chart points within a series inside a chart.
     */
    class ChartPointsCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Excel.ChartPoint>;
        /**
         *
         * Returns the number of chart points in the collection. Read-only.
         */
        count: number;
        /**
         *
         * Retrieve a point based on its position within the series.
         *
         * @param index Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Excel.ChartPoint;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartPointsCollection;
    }
    /**
     *
     * Represents a point of a series in a chart.
     */
    class ChartPoint extends OfficeExtension.ClientObject {
        private m_format;
        private m_value;
        /**
         *
         * Encapsulates the format properties chart point. Read-only.
         */
        format: Excel.ChartPointFormat;
        /**
         *
         * Returns the value of a chart point. Read-only.
         */
        value: any;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartPoint;
    }
    /**
     *
     * Represents formatting object for chart points.
     */
    class ChartPointFormat extends OfficeExtension.ClientObject {
        private m_fill;
        /**
         *
         * Represents the fill format of a chart, which includes background formating information. Read-only.
         */
        fill: Excel.ChartFill;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartPointFormat;
    }
    /**
     *
     * Represents the chart axes.
     */
    class ChartAxes extends OfficeExtension.ClientObject {
        private m_categoryAxis;
        private m_seriesAxis;
        private m_valueAxis;
        /**
         *
         * Represents the category axis in a chart. Read-only.
         */
        categoryAxis: Excel.ChartAxis;
        /**
         *
         * Represents the series axis of a 3-dimensional chart. Read-only.
         */
        seriesAxis: Excel.ChartAxis;
        /**
         *
         * Represents the value axis in an axis. Read-only.
         */
        valueAxis: Excel.ChartAxis;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartAxes;
    }
    /**
     *
     * Represents a single axis in a chart.
     */
    class ChartAxis extends OfficeExtension.ClientObject {
        private m_format;
        private m_majorGridlines;
        private m_majorUnit;
        private m_maximum;
        private m_minimum;
        private m_minorGridlines;
        private m_minorUnit;
        private m_title;
        /**
         *
         * Represents the formatting of a chart object, which includes line and font formatting. Read-only.
         */
        format: Excel.ChartAxisFormat;
        /**
         *
         * Returns a gridlines object that represents the major gridlines for the specified axis. Read-only.
         */
        majorGridlines: Excel.ChartGridlines;
        /**
         *
         * Returns a Gridlines object that represents the minor gridlines for the specified axis. Read-only.
         */
        minorGridlines: Excel.ChartGridlines;
        /**
         *
         * Represents the axis title. Read-only.
         */
        title: Excel.ChartAxisTitle;
        /**
         *
         * Represents the interval between two major tick marks. Can be set to a numeric value or an empty string.  The returned value is always a number.
         */
        majorUnit: any;
        /**
         *
         * Represents the maximum value on the value axis.  Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.
         */
        maximum: any;
        /**
         *
         * Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.
         */
        minimum: any;
        /**
         *
         * Represents the interval between two minor tick marks. "Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
         */
        minorUnit: any;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartAxis;
    }
    /**
     *
     * Encapsulates the format properties for the chart axis.
     */
    class ChartAxisFormat extends OfficeExtension.ClientObject {
        private m_font;
        private m_line;
        /**
         *
         * Represents the font attributes (font name, font size, color, etc.) for a chart axis element. Read-only.
         */
        font: Excel.ChartFont;
        /**
         *
         * Represents chart line formatting. Read-only.
         */
        line: Excel.ChartLineFormat;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartAxisFormat;
    }
    /**
     *
     * Represents the title of a chart axis.
     */
    class ChartAxisTitle extends OfficeExtension.ClientObject {
        private m_format;
        private m_text;
        private m_visible;
        /**
         *
         * Represents the formatting of chart axis title. Read-only.
         */
        format: Excel.ChartAxisTitleFormat;
        /**
         *
         * Represents the axis title.
         */
        text: string;
        /**
         *
         * A boolean that specifies the visibility of an axis title.
         */
        visible: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartAxisTitle;
    }
    /**
     *
     * Represents the chart axis title formatting.
     */
    class ChartAxisTitleFormat extends OfficeExtension.ClientObject {
        private m_font;
        /**
         *
         * Represents the font attributes, such as font name, font size, color, etc. of chart axis title object. Read-only.
         */
        font: Excel.ChartFont;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartAxisTitleFormat;
    }
    /**
     *
     * Represents a collection of all the data labels on a chart point.
     */
    class ChartDataLabels extends OfficeExtension.ClientObject {
        private m_format;
        private m_position;
        private m_separator;
        private m_showBubbleSize;
        private m_showCategoryName;
        private m_showLegendKey;
        private m_showPercentage;
        private m_showSeriesName;
        private m_showValue;
        /**
         *
         * Represents the format of chart data labels, which includes fill and font formatting. Read-only.
         */
        format: Excel.ChartDataLabelFormat;
        /**
         *
         * DataLabelPosition value that represents the position of the data label. See Excel.ChartDataLabelPosition for details.
         */
        position: string;
        /**
         *
         * String representing the separator used for the data labels on a chart.
         */
        separator: string;
        /**
         *
         * Boolean value representing if the data label bubble size is visible or not.
         */
        showBubbleSize: boolean;
        /**
         *
         * Boolean value representing if the data label category name is visible or not.
         */
        showCategoryName: boolean;
        /**
         *
         * Boolean value representing if the data label legend key is visible or not.
         */
        showLegendKey: boolean;
        /**
         *
         * Boolean value representing if the data label percentage is visible or not.
         */
        showPercentage: boolean;
        /**
         *
         * Boolean value representing if the data label series name is visible or not.
         */
        showSeriesName: boolean;
        /**
         *
         * Boolean value representing if the data label value is visible or not.
         */
        showValue: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartDataLabels;
    }
    /**
     *
     * Encapsulates the format properties for the chart data labels.
     */
    class ChartDataLabelFormat extends OfficeExtension.ClientObject {
        private m_fill;
        private m_font;
        /**
         *
         * Represents the fill format of the current chart data label. Read-only.
         */
        fill: Excel.ChartFill;
        /**
         *
         * Represents the font attributes (font name, font size, color, etc.) for a chart data label. Read-only.
         */
        font: Excel.ChartFont;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartDataLabelFormat;
    }
    /**
     *
     * Represents major or minor gridlines on a chart axis.
     */
    class ChartGridlines extends OfficeExtension.ClientObject {
        private m_format;
        private m_visible;
        /**
         *
         * Represents the formatting of chart gridlines. Read-only.
         */
        format: Excel.ChartGridlinesFormat;
        /**
         *
         * Boolean value representing if the axis gridlines are visible or not.
         */
        visible: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartGridlines;
    }
    /**
     *
     * Encapsulates the format properties for chart gridlines.
     */
    class ChartGridlinesFormat extends OfficeExtension.ClientObject {
        private m_line;
        /**
         *
         * Represents chart line formatting. Read-only.
         */
        line: Excel.ChartLineFormat;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartGridlinesFormat;
    }
    /**
     *
     * Represents the legend in a chart.
     */
    class ChartLegend extends OfficeExtension.ClientObject {
        private m_format;
        private m_overlay;
        private m_position;
        private m_visible;
        /**
         *
         * Represents the formatting of a chart legend, which includes fill and font formatting. Read-only.
         */
        format: Excel.ChartLegendFormat;
        /**
         *
         * Boolean value for whether the chart legend should overlap with the main body of the chart.
         */
        overlay: boolean;
        /**
         *
         * Represents the position of the legend on the chart. See Excel.ChartLegendPosition for details.
         */
        position: string;
        /**
         *
         * A boolean value the represents the visibility of a ChartLegend object.
         */
        visible: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartLegend;
    }
    /**
     *
     * Encapsulates the format properties of a chart legend.
     */
    class ChartLegendFormat extends OfficeExtension.ClientObject {
        private m_fill;
        private m_font;
        /**
         *
         * Represents the fill format of an object, which includes background formating information. Read-only.
         */
        fill: Excel.ChartFill;
        /**
         *
         * Represents the font attributes such as font name, font size, color, etc. of a chart legend. Read-only.
         */
        font: Excel.ChartFont;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartLegendFormat;
    }
    /**
     *
     * Represents a chart title object of a chart.
     */
    class ChartTitle extends OfficeExtension.ClientObject {
        private m_format;
        private m_overlay;
        private m_text;
        private m_visible;
        /**
         *
         * Represents the formatting of a chart title, which includes fill and font formatting. Read-only.
         */
        format: Excel.ChartTitleFormat;
        /**
         *
         * Boolean value representing if the chart title will overlay the chart or not.
         */
        overlay: boolean;
        /**
         *
         * Represents the title text of a chart.
         */
        text: string;
        /**
         *
         * A boolean value the represents the visibility of a chart title object.
         */
        visible: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartTitle;
    }
    /**
     *
     * Provides access to the office art formatting for chart title.
     */
    class ChartTitleFormat extends OfficeExtension.ClientObject {
        private m_fill;
        private m_font;
        /**
         *
         * Represents the fill format of an object, which includes background formating information. Read-only.
         */
        fill: Excel.ChartFill;
        /**
         *
         * Represents the font attributes (font name, font size, color, etc.) for an object. Read-only.
         */
        font: Excel.ChartFont;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartTitleFormat;
    }
    /**
     *
     * Represents the fill formatting for a chart element.
     */
    class ChartFill extends OfficeExtension.ClientObject {
        /**
         *
         * Clear the fill color of a chart element.
         *
         */
        clear(): void;
        /**
         *
         * Sets the fill formatting of a chart element to a uniform color.
         *
         * @param color HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").
         */
        setSolidColor(color: string): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartFill;
    }
    /**
     *
     * Enapsulates the formatting options for line elements.
     */
    class ChartLineFormat extends OfficeExtension.ClientObject {
        private m_color;
        /**
         *
         * HTML color code representing the color of lines in the chart.
         */
        color: string;
        /**
         *
         * Clear the line format of a chart element.
         *
         */
        clear(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartLineFormat;
    }
    /**
     *
     * This object represents the font attributes (font name, font size, color, etc.) for a chart object.
     */
    class ChartFont extends OfficeExtension.ClientObject {
        private m_bold;
        private m_color;
        private m_italic;
        private m_name;
        private m_size;
        private m_underline;
        /**
         *
         * Represents the bold status of font.
         */
        bold: boolean;
        /**
         *
         * HTML color code representation of the text color. E.g. #FF0000 represents Red.
         */
        color: string;
        /**
         *
         * Represents the italic status of the font.
         */
        italic: boolean;
        /**
         *
         * Font name (e.g. "Calibri")
         */
        name: string;
        /**
         *
         * Size of the font (e.g. 11)
         */
        size: number;
        /**
         *
         * Type of underline applied to the font. See Excel.ChartUnderlineStyle for details.
         */
        underline: string;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Excel.ChartFont;
    }
    module BindingType {
        var range: string;
        var table: string;
        var text: string;
    }
    module BorderIndex {
        var edgeTop: string;
        var edgeBottom: string;
        var edgeLeft: string;
        var edgeRight: string;
        var insideVertical: string;
        var insideHorizontal: string;
        var diagonalDown: string;
        var diagonalUp: string;
    }
    module BorderLineStyle {
        var none: string;
        var continuous: string;
        var dash: string;
        var dashDot: string;
        var dashDotDot: string;
        var dot: string;
        var double: string;
        var slantDashDot: string;
    }
    module BorderWeight {
        var hairline: string;
        var thin: string;
        var medium: string;
        var thick: string;
    }
    module CalculationMode {
        var automatic: string;
        var automaticExceptTables: string;
        var manual: string;
    }
    module CalculationType {
        var recalculate: string;
        var full: string;
        var fullRebuild: string;
    }
    module ClearApplyTo {
        var all: string;
        var formats: string;
        var contents: string;
    }
    module ChartDataLabelPosition {
        var invalid: string;
        var none: string;
        var center: string;
        var insideEnd: string;
        var insideBase: string;
        var outsideEnd: string;
        var left: string;
        var right: string;
        var top: string;
        var bottom: string;
        var bestFit: string;
        var callout: string;
    }
    module ChartLegendPosition {
        var invalid: string;
        var top: string;
        var bottom: string;
        var left: string;
        var right: string;
        var corner: string;
        var custom: string;
    }
    module ChartSeriesBy {
        var auto: string;
        var columns: string;
        var rows: string;
    }
    module ChartType {
        var invalid: string;
        var columnClustered: string;
        var columnStacked: string;
        var columnStacked100: string;
        var _3DColumnClustered: string;
        var _3DColumnStacked: string;
        var _3DColumnStacked100: string;
        var barClustered: string;
        var barStacked: string;
        var barStacked100: string;
        var _3DBarClustered: string;
        var _3DBarStacked: string;
        var _3DBarStacked100: string;
        var lineStacked: string;
        var lineStacked100: string;
        var lineMarkers: string;
        var lineMarkersStacked: string;
        var lineMarkersStacked100: string;
        var pieOfPie: string;
        var pieExploded: string;
        var _3DPieExploded: string;
        var barOfPie: string;
        var xyscatterSmooth: string;
        var xyscatterSmoothNoMarkers: string;
        var xyscatterLines: string;
        var xyscatterLinesNoMarkers: string;
        var areaStacked: string;
        var areaStacked100: string;
        var _3DAreaStacked: string;
        var _3DAreaStacked100: string;
        var doughnutExploded: string;
        var radarMarkers: string;
        var radarFilled: string;
        var surface: string;
        var surfaceWireframe: string;
        var surfaceTopView: string;
        var surfaceTopViewWireframe: string;
        var bubble: string;
        var bubble3DEffect: string;
        var stockHLC: string;
        var stockOHLC: string;
        var stockVHLC: string;
        var stockVOHLC: string;
        var cylinderColClustered: string;
        var cylinderColStacked: string;
        var cylinderColStacked100: string;
        var cylinderBarClustered: string;
        var cylinderBarStacked: string;
        var cylinderBarStacked100: string;
        var cylinderCol: string;
        var coneColClustered: string;
        var coneColStacked: string;
        var coneColStacked100: string;
        var coneBarClustered: string;
        var coneBarStacked: string;
        var coneBarStacked100: string;
        var coneCol: string;
        var pyramidColClustered: string;
        var pyramidColStacked: string;
        var pyramidColStacked100: string;
        var pyramidBarClustered: string;
        var pyramidBarStacked: string;
        var pyramidBarStacked100: string;
        var pyramidCol: string;
        var _3DColumn: string;
        var line: string;
        var _3DLine: string;
        var _3DPie: string;
        var pie: string;
        var xyscatter: string;
        var _3DArea: string;
        var area: string;
        var doughnut: string;
        var radar: string;
    }
    module ChartUnderlineStyle {
        var none: string;
        var single: string;
    }
    module DeleteShiftDirection {
        var up: string;
        var left: string;
    }
    module HorizontalAlignment {
        var general: string;
        var left: string;
        var center: string;
        var right: string;
        var fill: string;
        var justify: string;
        var centerAcrossSelection: string;
        var distributed: string;
    }
    module InsertShiftDirection {
        var down: string;
        var right: string;
    }
    module NamedItemType {
        var string: string;
        var integer: string;
        var double: string;
        var boolean: string;
        var range: string;
    }
    module RangeUnderlineStyle {
        var none: string;
        var single: string;
        var double: string;
        var singleAccountant: string;
        var doubleAccountant: string;
    }
    module SheetVisibility {
        var visible: string;
        var hidden: string;
        var veryHidden: string;
    }
    module RangeValueType {
        var unknown: string;
        var empty: string;
        var string: string;
        var integer: string;
        var double: string;
        var boolean: string;
        var error: string;
    }
    module VerticalAlignment {
        var top: string;
        var center: string;
        var bottom: string;
        var justify: string;
        var distributed: string;
    }
    module ErrorCodes {
        var accessDenied: string;
        var generalException: string;
        var insertDeleteConflict: string;
        var invalidArgument: string;
        var invalidBinding: string;
        var invalidOperation: string;
        var invalidReference: string;
        var invalidSelection: string;
        var itemAlreadyExists: string;
        var itemNotFound: string;
        var notImplemented: string;
        var unsupportedOperation: string;
    }
}
declare module Excel {
    class RequestContext extends OfficeExtension.ClientRequestContext {
        private m_workbook;
        constructor(url?: string);
        workbook: Workbook;
    }
}
