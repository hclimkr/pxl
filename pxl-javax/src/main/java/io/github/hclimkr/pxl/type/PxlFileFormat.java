package io.github.hclimkr.pxl.type;

import com.github.pjfanning.xlsx.impl.StreamingWorkbook;
import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Physical spreadsheet file format — what the bytes are — each carrying its filename extension, MIME content
 * type, and the maximum sheet/row/column counts applied on import and export.
 * <p>
 * This is the format axis, not the writer axis. Which POI implementation writes an Excel file is
 * {@link PxlExcelEngine}, and two engines write the same format: both {@code XSSF} and {@code SXSSF} produce
 * {@link #XLSX}. A workbook class declares the engine through {@code @PxlWorkbook(exportExcelEngine = ...)},
 * which {@link PxlExcelEngine#fromWorkbookObject(Class)} reads back.
 * <p>
 * {@link #fromPoiWorkbook(Workbook)} recovers the format an open POI workbook holds. It is a plain lookup — it
 * throws nothing and never returns {@code null}, falling back to
 * {@link PxlConstants#DEFAULT_EXPORT_FILE_FORMAT}.
 */
@Getter
@AllArgsConstructor
public enum PxlFileFormat {

    // Horrible SpreadSheet Format (Excel '97)
    /**
     * Legacy binary {@code .xls} format (Excel '97), written by {@link PxlExcelEngine#HSSF} and bounded by the
     * Excel 97 sheet limits.
     */
    XLS(
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
     * OOXML {@code .xlsx} format (Excel 2007), the default export format, bounded by the Excel 2007 sheet
     * limits. Both {@link PxlExcelEngine#XSSF} and {@link PxlExcelEngine#SXSSF} write it — they differ in how
     * much of the workbook they hold in memory, not in what they produce.
     */
    XLSX(
            PxlConstants.FILENAME_EXTENSION_XLSX,
            PxlConstants.CONTENT_TYPE_MICROSOFT_XLSX,

            PxlConstants.IMPORT_MAX_NUMBER_OF_EXCEL_SHEETS,
            SpreadsheetVersion.EXCEL2007.getMaxRows(),
            SpreadsheetVersion.EXCEL2007.getMaxColumns(),

            PxlConstants.EXPORT_MAX_NUMBER_OF_EXCEL_SHEETS,
            SpreadsheetVersion.EXCEL2007.getMaxRows(),
            SpreadsheetVersion.EXCEL2007.getMaxColumns()
    ),

    // Comma Separated Values
    /**
     * Comma-separated values ({@code .csv}) format, bounded by the dedicated CSV sheet/row/column limits.
     * No {@link PxlExcelEngine} produces it — it is plain text rather than a POI workbook.
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

    /**
     * Resolves the file format an open POI workbook holds, from its implementation type.
     * <p>
     * {@link HSSFWorkbook} maps to {@link #XLS}; {@link XSSFWorkbook}, {@link SXSSFWorkbook} and the streaming
     * reader's {@link StreamingWorkbook} all map to {@link #XLSX}, because all three sit on the same OOXML
     * container and differ only in how they read or write it. Which of them it is, is the engine question that
     * {@link PxlExcelEngine#fromPoiWorkbook(Workbook)} answers. {@link #CSV} is never returned, as no POI
     * workbook represents it.
     * <p>
     * This is a plain lookup: it throws nothing and never returns {@code null} — a {@code null} argument or an
     * unrecognized workbook type falls back to {@link PxlConstants#DEFAULT_EXPORT_FILE_FORMAT}.
     *
     * @param poiWorkbook the open POI workbook to inspect (may be {@code null})
     * @return the file format matching the workbook implementation type, or
     * {@link PxlConstants#DEFAULT_EXPORT_FILE_FORMAT} if {@code poiWorkbook} is {@code null} or its
     * implementation type matches no file format
     */
    public static PxlFileFormat fromPoiWorkbook(@Nullable final Workbook poiWorkbook) {

        if (poiWorkbook instanceof HSSFWorkbook) {
            return XLS;
        }

        // SXSSFWorkbook wraps an XSSFWorkbook instead of extending it, so the order of these checks is not load-bearing.
        // The streaming reader opens the same OOXML container as well, so all three report XLSX.
        if (poiWorkbook instanceof XSSFWorkbook
                || poiWorkbook instanceof SXSSFWorkbook
                || poiWorkbook instanceof StreamingWorkbook) {
            return XLSX;
        }

        return PxlConstants.DEFAULT_EXPORT_FILE_FORMAT;
    }

}
