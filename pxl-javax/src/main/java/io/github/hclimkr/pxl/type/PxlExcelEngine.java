package io.github.hclimkr.pxl.type;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Objects;
import java.util.Optional;

/**
 * POI implementation that writes an Excel workbook on export - the engine axis, as opposed to the physical
 * format axis of {@link PxlFileFormat}.
 * <p>
 * Each engine writes exactly one physical format, and two of them write the same one: {@link #HSSF} produces
 * {@link PxlFileFormat#XLS} while both {@link #XSSF} and {@link #SXSSF} produce {@link PxlFileFormat#XLSX}, the
 * latter streaming its rows to keep the heap independent of the row count. The sheet/row/column limits therefore
 * belong to the file format and are reached through {@code getFileFormat()}.
 * <p>
 * There is deliberately no constant for CSV: CSV is a physical format with no POI workbook behind it, so it is
 * written through a dedicated CSV path rather than by an engine. That absence is what keeps an engine-typed
 * declaration from naming a format it cannot produce.
 */
@Getter
@AllArgsConstructor
public enum PxlExcelEngine {

    // Horrible SpreadSheet Format (Excel '97)
    /**
     * Legacy binary writer (POI HSSF) producing {@link PxlFileFormat#XLS}, bounded by the Excel 97 sheet limits.
     */
    HSSF(PxlFileFormat.XLS),

    // XML SpreadSheet Format (Excel 2007)
    /**
     * OOXML writer (POI XSSF) producing {@link PxlFileFormat#XLSX}; the default engine, which builds the whole
     * workbook in memory before writing.
     */
    XSSF(PxlFileFormat.XLSX),

    // Streaming XML SpreadSheet Format
    /**
     * Streaming OOXML writer (POI SXSSF) producing the same {@link PxlFileFormat#XLSX} as {@link #XSSF}, but
     * keeping only a sliding window of rows in memory and spilling the rest to temporary files.
     */
    SXSSF(PxlFileFormat.XLSX),
    ;

    /**
     * Physical file format this engine writes; the source of the sheet/row/column limits.
     */
    private final PxlFileFormat fileFormat;

    /**
     * Finds and returns the export engine declared by a workbook class through
     * {@link PxlWorkbook#exportExcelEngine()}.
     * <p>
     * This is a plain lookup: it throws nothing and never returns {@code null} - a {@code null} class, a class
     * without {@code @PxlWorkbook}, or an annotation left at its default all yield
     * {@link PxlConstants#DEFAULT_EXPORT_EXCEL_ENGINE}.
     *
     * @param workbookClass the workbook class to inspect (may be {@code null})
     * @return the declared export engine, or {@link PxlConstants#DEFAULT_EXPORT_EXCEL_ENGINE} if absent
     */
    public static PxlExcelEngine fromWorkbookObject(@Nullable final Class<?> workbookClass) {

        if (Objects.isNull(workbookClass)) {
            return PxlConstants.DEFAULT_EXPORT_EXCEL_ENGINE;
        }

        final PxlWorkbook workbookAnnotation = workbookClass.getAnnotation(PxlWorkbook.class);

        return Optional.ofNullable(workbookAnnotation)
                .map(PxlWorkbook::exportExcelEngine)
                .orElse(PxlConstants.DEFAULT_EXPORT_EXCEL_ENGINE);
    }

    /**
     * Resolves the export engine of an open POI workbook from its implementation type.
     * <p>
     * Only the three writer implementations map to an engine - {@link HSSFWorkbook} to {@link #HSSF},
     * {@link SXSSFWorkbook} to {@link #SXSSF}, and {@link XSSFWorkbook} to {@link #XSSF}. The streaming reader's
     * workbook is a reader, not a writer, and therefore has no engine; ask
     * {@link PxlFileFormat#fromPoiWorkbook(Workbook)} what format such a workbook holds instead.
     * <p>
     * This is a plain lookup: it throws nothing and never returns {@code null} - a {@code null} argument or a
     * workbook that is not one of the three writers falls back to {@link PxlConstants#DEFAULT_EXPORT_EXCEL_ENGINE}.
     *
     * @param poiWorkbook the open POI workbook to inspect (may be {@code null})
     * @return the engine matching the workbook implementation type, or
     * {@link PxlConstants#DEFAULT_EXPORT_EXCEL_ENGINE} if {@code poiWorkbook} is {@code null} or matches no engine
     */
    public static PxlExcelEngine fromPoiWorkbook(@Nullable final Workbook poiWorkbook) {

        if (poiWorkbook instanceof HSSFWorkbook) {
            return HSSF;
        }

        // SXSSFWorkbook wraps an XSSFWorkbook instead of extending it, so the order of these checks is not load-bearing.
        if (poiWorkbook instanceof SXSSFWorkbook) {
            return SXSSF;
        }

        if (poiWorkbook instanceof XSSFWorkbook) {
            return XSSF;
        }

        return PxlConstants.DEFAULT_EXPORT_EXCEL_ENGINE;
    }

}
