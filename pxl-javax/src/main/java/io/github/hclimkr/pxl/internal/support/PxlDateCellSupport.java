package io.github.hclimkr.pxl.internal.support;

import io.github.hclimkr.pxl.internal.meta.PxlExportColumnMeta;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Date;
import java.util.Objects;

/**
 * Shared helper for exporting date/time values as Numeric (Excel date serial) cells rather than as strings.
 * <p>
 * Each date/time codec writes the value through this helper when no pattern or masking is specified ({@link PxlExportColumnMeta#isExportedToString()} is false).
 * In this case the cell gets a date display-format style (inherited from the column-common data style), so that
 * the value itself is numeric but is displayed on screen as a date/time, and Excel's date sorting, filtering, and functions work correctly.
 */
public final class PxlDateCellSupport {

    /**
     * Prevents instantiation.
     */
    private PxlDateCellSupport() {

        throw new AssertionError("no instances of this class");
    }

    // Number of nanoseconds in a day in Excel time (one day = 1.0). Used to convert a time (LocalTime/OffsetTime) into a fraction.
    private static final double NANOS_PER_DAY = 86_400_000_000_000.0;

    // Applies a date display-format style (inheriting the base data style) to the cell. When creation fails, proceeds without a style.

    /**
     * Applies a date display-format style (inheriting the column-common data style) to the cell. When the style cannot be
     * created it is left unset and the cell keeps its existing style.
     *
     * @param cell            the target cell
     * @param columnMeta      the column metadata (supplies/creates the date-formatted style)
     * @param excelFormatCode the Excel number-format code used to display the value as a date
     */
    private static void applyDateFormat(final Cell cell,
                                        final PxlExportColumnMeta columnMeta,
                                        final String excelFormatCode) {

        final CellStyle dateFormattedCellStyle = columnMeta.getWorkbookMeta().getOrCreateDateFormattedCellStyle(cell, excelFormatCode);
        if (Objects.nonNull(dateFormattedCellStyle)) {
            cell.setCellStyle(dateFormattedCellStyle);
        }
    }

    /**
     * Writes a legacy {@link Date} as a numeric (Excel date-serial) cell and applies the date display-format style.
     *
     * @param cell            the target cell
     * @param value           the date value
     * @param columnMeta      the column metadata (supplies/creates the date-formatted style)
     * @param excelFormatCode the Excel number-format code used to display the value as a date
     */
    public static void writeNumericCell(final Cell cell,
                                        final Date value,
                                        final PxlExportColumnMeta columnMeta,
                                        final String excelFormatCode) {

        applyDateFormat(cell, columnMeta, excelFormatCode);
        cell.setCellValue(value);
    }

    /**
     * Writes a {@link LocalDate} as a numeric (Excel date-serial) cell and applies the date display-format style.
     *
     * @param cell            the target cell
     * @param value           the date value
     * @param columnMeta      the column metadata (supplies/creates the date-formatted style)
     * @param excelFormatCode the Excel number-format code used to display the value as a date
     */
    public static void writeNumericCell(final Cell cell,
                                        final LocalDate value,
                                        final PxlExportColumnMeta columnMeta,
                                        final String excelFormatCode) {

        applyDateFormat(cell, columnMeta, excelFormatCode);
        cell.setCellValue(value);
    }

    /**
     * Writes a {@link LocalDateTime} as a numeric (Excel date-serial) cell and applies the date display-format style.
     *
     * @param cell            the target cell
     * @param value           the date/time value
     * @param columnMeta      the column metadata (supplies/creates the date-formatted style)
     * @param excelFormatCode the Excel number-format code used to display the value as a date/time
     */
    public static void writeNumericCell(final Cell cell,
                                        final LocalDateTime value,
                                        final PxlExportColumnMeta columnMeta,
                                        final String excelFormatCode) {

        applyDateFormat(cell, columnMeta, excelFormatCode);
        cell.setCellValue(value);
    }

    // POI has no time-specific setCellValue overload, so write the value converted to a fraction of a day (0.0–1.0).

    /**
     * Writes a {@link LocalTime} as a numeric cell, converting it to the Excel fraction-of-a-day representation
     * (0.0–1.0), since POI has no time-specific {@code setCellValue} overload. Applies the date/time display-format style.
     *
     * @param cell            the target cell
     * @param value           the time value
     * @param columnMeta      the column metadata (supplies/creates the date-formatted style)
     * @param excelFormatCode the Excel number-format code used to display the value as a time
     */
    public static void writeNumericTimeCell(final Cell cell,
                                            final LocalTime value,
                                            final PxlExportColumnMeta columnMeta,
                                            final String excelFormatCode) {

        applyDateFormat(cell, columnMeta, excelFormatCode);
        cell.setCellValue((double) value.toNanoOfDay() / NANOS_PER_DAY);
    }

}
