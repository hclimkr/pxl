package io.github.hclimkr.pxl;

import com.github.pjfanning.xlsx.impl.StreamingWorkbook;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Objects;
import java.util.Optional;

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

    /**
     * Resolves the file format of an open POI workbook from its implementation type.
     * <p>
     * {@link HSSFWorkbook} maps to {@link #HSSF} and {@link SXSSFWorkbook} to {@link #SXSSF}, while both
     * {@link XSSFWorkbook} and the streaming reader's {@code StreamingWorkbook} map to {@link #XSSF} — a
     * streamed workbook is read from the same OOXML ({@code .xlsx}) container, and {@link #SXSSF} denotes the
     * streaming <em>export</em> workbook only. {@link #CSV} is never returned, as no POI workbook represents it.
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
            return HSSF;
        }

        // SXSSFWorkbook wraps an XSSFWorkbook instead of extending it, so the order of these checks is not load-bearing.
        if (poiWorkbook instanceof SXSSFWorkbook) {
            return SXSSF;
        }

        // The streaming reader opens the same OOXML container as XSSF, so it reports the XSSF format.
        if (poiWorkbook instanceof XSSFWorkbook || poiWorkbook instanceof StreamingWorkbook) {
            return XSSF;
        }

        return PxlConstants.DEFAULT_EXPORT_FILE_FORMAT;
    }

    /**
     * Finds and returns the export file format declared by a workbook class through
     * {@link PxlWorkbook#exportFileFormat()}.
     * <p>
     * Like {@link #fromPoiWorkbook(Workbook)}, this is a plain lookup: it throws nothing and never returns
     * {@code null} — a {@code null} class, a class without {@code @PxlWorkbook}, or an annotation left at its
     * default all yield {@link PxlConstants#DEFAULT_EXPORT_FILE_FORMAT}.
     *
     * @param workbookClass the workbook class to inspect (may be {@code null})
     * @return the declared export file format, or {@link PxlConstants#DEFAULT_EXPORT_FILE_FORMAT} if absent
     */
    public static PxlFileFormat fromWorkbookObject(@Nullable final Class<?> workbookClass) {

        if (Objects.isNull(workbookClass)) {
            return PxlConstants.DEFAULT_EXPORT_FILE_FORMAT;
        }

        final PxlWorkbook workbookAnnotation = workbookClass.getAnnotation(PxlWorkbook.class);

        return Optional.ofNullable(workbookAnnotation)
                .map(PxlWorkbook::exportFileFormat)
                .orElse(PxlConstants.DEFAULT_EXPORT_FILE_FORMAT);
    }

}
