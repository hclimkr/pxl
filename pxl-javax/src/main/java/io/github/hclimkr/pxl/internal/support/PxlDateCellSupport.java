package io.github.hclimkr.pxl.internal.support;

import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlExportColumnMeta;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Date;
import java.util.Objects;

/**
 * Shared helper for the Numeric (Excel date serial) cell of a date/time column, in both directions.
 * <p>
 * On export, each date/time codec writes the value through this helper when no pattern or masking is specified ({@link PxlExportColumnMeta#isExportedToString()} is false).
 * In this case the cell gets a date display-format style (inherited from the column-common data style), so that
 * the value itself is numeric but is displayed on screen as a date/time, and Excel's date sorting, filtering, and functions work correctly.
 * <p>
 * On import, the codecs read a numeric cell through {@link #readNumericCellAsLocalDateTime} / {@link #readNumericCellAsJavaDate},
 * which turn a serial number that is no Excel date into a diagnostic naming the value.
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

    // POI has no time-specific setCellValue overload, so write the value converted to a fraction of a day (0.0-1.0).

    /**
     * Writes a {@link LocalTime} as a numeric cell, converting it to the Excel fraction-of-a-day representation
     * (0.0-1.0), since POI has no time-specific {@code setCellValue} overload. Applies the date/time display-format style.
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

    /**
     * Reads a numeric cell as a {@link LocalDateTime}, rejecting a serial number that is no Excel date. (import)
     * <p>
     * A date-formatted cell is read through POI's own date conversion, any other numeric cell as a raw Excel date
     * serial. POI answers {@code null} for a serial outside the Excel date range - a negative one, for instance -
     * so the value is checked here instead of being dereferenced by each caller, which would surface the failure as
     * a message-less {@link NullPointerException} that names neither the cell nor the value.
     *
     * @param cell     the numeric cell to read
     * @param typeName the name of the target type, used in the diagnostic message
     * @return the cell value as a {@link LocalDateTime}
     * @throws PxlCellCodecException if the cell holds a serial number that is not a valid Excel date
     */
    public static LocalDateTime readNumericCellAsLocalDateTime(final Cell cell,
                                                               final String typeName)
            throws PxlCellCodecException {

        final LocalDateTime localDateTimeValue = DateUtil.isCellDateFormatted(cell)
                ? cell.getLocalDateTimeCellValue()
                : DateUtil.getLocalDateTime(cell.getNumericCellValue());

        if (Objects.isNull(localDateTimeValue)) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(cell), typeName));
        }

        return localDateTimeValue;
    }

    /**
     * Reads a numeric cell as a {@link Date}, rejecting a serial number that is no Excel date. (import)
     * <p>
     * The {@link Date} counterpart of {@link #readNumericCellAsLocalDateTime}: POI answers {@code null} for
     * the same out-of-range serials, and the Date codec assigned that {@code null} straight to the field, leaving an
     * invalid cell to bind as no value at all instead of being reported.
     *
     * @param cell     the numeric cell to read
     * @param typeName the name of the target type, used in the diagnostic message
     * @return the cell value as a {@link Date}
     * @throws PxlCellCodecException if the cell holds a serial number that is not a valid Excel date
     */
    public static Date readNumericCellAsJavaDate(final Cell cell,
                                                 final String typeName)
            throws PxlCellCodecException {

        final Date dateValue = DateUtil.isCellDateFormatted(cell)
                ? cell.getDateCellValue()
                : DateUtil.getJavaDate(cell.getNumericCellValue());

        if (Objects.isNull(dateValue)) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(cell), typeName));
        }

        return dateValue;
    }

}
