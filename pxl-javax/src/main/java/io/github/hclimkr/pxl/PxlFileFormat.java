package io.github.hclimkr.pxl;

import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.ss.SpreadsheetVersion;

/**
 * Excel/CSV file format, each carrying its filename extension, MIME content type, and the maximum
 * sheet/row/column counts applied on import and export.
 */
@Getter
@AllArgsConstructor
public enum PxlFileFormat {

    // Horrible SpreadSheet Format (Excel '97)
    /**
     * Legacy binary {@code .xls} format (Excel '97; POI HSSF), bounded by the Excel 97 sheet limits.
     */
    HSSF(
            PxlConstants.FILENAME_EXTENSION_XLS,
            PxlConstants.CONTENT_TYPE_MICROSOFT_XLS,

            PxlConstants.IMPORT_MAX_NUMBER_OF_EXCEL_SHEETS,
            SpreadsheetVersion.EXCEL97.getMaxRows(),
            SpreadsheetVersion.EXCEL97.getMaxColumns(),

            PxlConstants.EXPORT_MAX_NUMBER_OF_EXCEL_SHEETS,
            SpreadsheetVersion.EXCEL97.getMaxRows(),
            SpreadsheetVersion.EXCEL97.getMaxColumns()
    ),

    // XML SpreadSheet Format (Excel 2007)
    /**
     * OOXML {@code .xlsx} format (Excel 2007; POI XSSF), the default export format, bounded by the
     * Excel 2007 sheet limits.
     */
    XSSF(
            PxlConstants.FILENAME_EXTENSION_XLSX,
            PxlConstants.CONTENT_TYPE_MICROSOFT_XLSX,

            PxlConstants.IMPORT_MAX_NUMBER_OF_EXCEL_SHEETS,
            SpreadsheetVersion.EXCEL2007.getMaxRows(),
            SpreadsheetVersion.EXCEL2007.getMaxColumns(),

            PxlConstants.EXPORT_MAX_NUMBER_OF_EXCEL_SHEETS,
            SpreadsheetVersion.EXCEL2007.getMaxRows(),
            SpreadsheetVersion.EXCEL2007.getMaxColumns()
    ),

    // Streaming XML SpreadSheet Format
    /**
     * Streaming OOXML {@code .xlsx} format (POI SXSSF) for low-memory export; shares the Excel 2007
     * sheet limits and the {@code .xlsx} extension/content type.
     */
    SXSSF(
            PxlConstants.FILENAME_EXTENSION_XLSX,
            PxlConstants.CONTENT_TYPE_MICROSOFT_XLSX,

            PxlConstants.IMPORT_MAX_NUMBER_OF_EXCEL_SHEETS,
            SpreadsheetVersion.EXCEL2007.getMaxRows(),
            SpreadsheetVersion.EXCEL2007.getMaxColumns(),

            PxlConstants.EXPORT_MAX_NUMBER_OF_EXCEL_SHEETS,
            SpreadsheetVersion.EXCEL2007.getMaxRows(),
            SpreadsheetVersion.EXCEL2007.getMaxColumns()
    ),

    // not supported yet
    // Comma Separated Values
    /**
     * Comma-separated values ({@code .csv}) format, bounded by the dedicated CSV sheet/row/column limits.
     */
    CSV(
            PxlConstants.FILENAME_EXTENSION_CSV,
            PxlConstants.CONTENT_TYPE_CSV,

            PxlConstants.IMPORT_MAX_NUMBER_OF_CSV_SHEETS,
            PxlConstants.IMPORT_MAX_NUMBER_OF_CSV_ROWS,
            PxlConstants.IMPORT_MAX_NUMBER_OF_CSV_COLUMNS,

            PxlConstants.EXPORT_MAX_NUMBER_OF_CSV_SHEETS,
            PxlConstants.EXPORT_MAX_NUMBER_OF_CSV_ROWS,
            PxlConstants.EXPORT_MAX_NUMBER_OF_CSV_COLUMNS
    ),
    ;

    /**
     * Filename extension for this format (without the dot), e.g. {@code "xlsx"}.
     */
    private final String filenameExtension;
    /**
     * MIME content type for this format.
     */
    private final String contentType;

    /**
     * Maximum number of sheets accepted on import.
     */
    private final int maxImportSheets;
    /**
     * Maximum number of rows accepted per sheet on import.
     */
    private final int maxImportRows;
    /**
     * Maximum number of columns accepted per sheet on import.
     */
    private final int maxImportColumns;

    /**
     * Maximum number of sheets allowed on export.
     */
    private final int maxExportSheets;
    /**
     * Maximum number of rows allowed per sheet on export.
     */
    private final int maxExportRows;
    /**
     * Maximum number of columns allowed per sheet on export.
     */
    private final int maxExportColumns;

}
