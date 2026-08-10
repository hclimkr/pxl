package io.github.hclimkr.pxl.internal.codec;

import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlExportColumnMeta;
import io.github.hclimkr.pxl.internal.meta.PxlImportColumnMeta;
import io.github.hclimkr.pxl.internal.support.PxlNumberSupport;
import io.github.hclimkr.pxl.internal.support.PxlTemporalAmountSupport;
import org.apache.commons.lang3.BooleanUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DurationFormatUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.time.LocalDateTime;
import java.time.Period;
import java.time.ZoneId;
import java.time.format.DateTimeParseException;
import java.util.Objects;
import java.util.Optional;

/**
 * Codec for {@link Period} column values - parses cells/strings into {@link Period} on import
 * and writes {@link Period} into cells on export.
 *
 * <p>NUMERIC and BOOLEAN cells are interpreted as a number of days. Strings are parsed by the column's
 * import pattern when set, otherwise by ISO-8601 {@link Period#parse}. Export uses the column's
 * export pattern (computed over {@code now .. now+period}) when set, otherwise the ISO-8601 form.
 */
final class PxlPeriodCodec {

    /**
     * Prevents instantiation.
     */
    private PxlPeriodCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses the given cell into a {@link Period}. NUMERIC cells are treated as a number of days
     * (range-checked to {@code int}); BOOLEAN cells map to 1 or 0 days; STRING cells are delegated to the
     * string parser; BLANK cells yield {@code null}.
     *
     * @param cell       the cell to read
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link Period}, or {@code null} when blank
     * @throws PxlCellCodecException if the cell type is unsupported or the numeric value is out of range
     */
    static Period parsePeriodValue(final Cell cell,
                                   final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Period periodValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                int intValue = PxlNumberSupport.requireWithinRange(numericValue, Integer.MIN_VALUE, Integer.MAX_VALUE, "Period").intValue();
                // Arbitrarily assume the value is given in (days).
                periodValue = Period.ofDays(intValue);
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                periodValue = parsePeriodValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                final boolean booleanCellValue = cell.getBooleanCellValue();
                // Arbitrarily assume the value is given in (days).
                periodValue = Period.ofDays(BooleanUtils.toInteger(booleanCellValue));
                break;

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return periodValue;
    }

    /**
     * Parses a string token into a {@link Period}. When the column's import pattern is set it is
     * tried first, falling back to ISO-8601 {@link Period#parse} on mismatch. The value is
     * trimmed when {@code importTrim} is set; a blank value yields {@code null}.
     *
     * @param s          the raw string token
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link Period}, or {@code null} when blank
     * @throws PxlCellCodecException if the value is not a valid period
     */
    static Period parsePeriodValue(final String s,
                                   final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        final PxlTemporalAmountSupport.CompiledTemporalPattern importTemporalPattern = columnMeta.getImportTemporalPatternCache();
        if (Objects.nonNull(importTemporalPattern)) {
            try {
                return PxlTemporalAmountSupport.parsePeriodByPattern(stringValue, importTemporalPattern);
            } catch (IllegalArgumentException ignored) {
                // On pattern mismatch, fall back to ISO-8601 parsing
            }
        }

        try {
            return Period.parse(stringValue);
        } catch (DateTimeParseException dateTimeParseException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "Period"), dateTimeParseException);
        }
    }

    /**
     * Writes the given value as a {@link Period} cell and returns the exported string. A {@link String}
     * source is parsed with ISO-8601 {@link Period#parse}; a {@link Period} source is
     * used directly. A {@code null} result blanks the cell; otherwise the value is formatted with the
     * column's export pattern when set, else its ISO-8601 form.
     *
     * @param cell       the target cell, or {@code null} to only compute the string
     * @param object     the source value (a {@link String} or {@link Period})
     * @param columnMeta the resolved export metadata for this column
     * @return the exported string, or {@code null} when blank
     * @throws PxlCellCodecException if the source is unsupported, the string is invalid, or the value is too large
     */
    static String buildPeriodCell(final Cell cell,
                                  final Object object,
                                  final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Period periodValue;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                periodValue = null;
            } else {
                try {
                    periodValue = Period.parse(stringValue);
                } catch (DateTimeParseException dateTimeParseException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(stringValue), "Period"), dateTimeParseException);
                }
            }
        } else if (object instanceof Period) {
            periodValue = (Period) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "Period"));
        }

        if (Objects.isNull(periodValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else {
            final String cellString = makePeriodExportString(periodValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(cellString));
            return cellString;
        }
    }

    /**
     * Renders the export string for a {@link Period}: when an export pattern is set the period is applied to the current
     * date-time and the resulting span is formatted with {@link DurationFormatUtils}, otherwise its {@link Period#toString()}
     * ISO-8601 form is used; then applies string-level export processing via {@link PxlStringCodec#makeExportString}.
     *
     * @param periodValue the value to render
     * @param columnMeta  resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the span overflows or the export pattern cannot be applied
     */
    private static String makePeriodExportString(final Period periodValue,
                                                 final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(periodValue)) {
            return null;
        }

        final String exportPattern = columnMeta.getExportPattern();

        try {
            String stringValue;

            if (StringUtils.isNotBlank(exportPattern)) {
                final LocalDateTime startDateTime = LocalDateTime.now();
                final LocalDateTime endDateTime = startDateTime.plus(periodValue);
                final long startMillis = startDateTime.atZone(ZoneId.systemDefault()).toInstant().toEpochMilli();
                final long endMillis = endDateTime.atZone(ZoneId.systemDefault()).toInstant().toEpochMilli();

                stringValue = DurationFormatUtils.formatPeriod(startMillis, endMillis, exportPattern);
            } else {
                // stringValue = DurationFormatUtils.formatPeriodISO(startMillis, endMillis);
                stringValue = periodValue.toString();
            }

            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } catch (ArithmeticException arithmeticException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_VALUE_TOO_LARGE, String.valueOf(periodValue), "Period"), arithmeticException);
        } catch (IllegalArgumentException illegalArgumentException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(periodValue), "Period"), illegalArgumentException);
        }
    }

}
