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

import java.time.Duration;
import java.time.format.DateTimeParseException;
import java.util.Objects;
import java.util.Optional;

/**
 * Codec for {@link Duration} column values — parses cells/strings into {@link Duration} on
 * import and writes {@link Duration} into cells on export.
 *
 * <p>NUMERIC and BOOLEAN cells are interpreted as a number of seconds. Strings are parsed by the column's
 * import pattern (via {@link DurationFormatUtils}) when set, otherwise by ISO-8601
 * {@link Duration#parse}. Export uses the column's export pattern when set, otherwise the
 * ISO-8601 form.
 */
final class PxlDurationCodec {

    /**
     * Prevents instantiation.
     */
    private PxlDurationCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses the given cell into a {@link Duration}. NUMERIC cells are treated as a number of
     * seconds (range-checked to {@code long}); BOOLEAN cells map to 1 or 0 seconds; STRING cells are
     * delegated to the string parser; BLANK cells yield {@code null}.
     *
     * @param cell       the cell to read
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link Duration}, or {@code null} when blank
     * @throws PxlCellCodecException if the cell type is unsupported or the numeric value is out of range
     */
    static Duration parseDurationValue(final Cell cell,
                                       final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Duration durationValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                long longValue = PxlNumberSupport.requireWithinRange(numericValue, Long.MIN_VALUE, Long.MAX_VALUE, "Duration").longValue();
                // Arbitrarily assume the value is given in (seconds).
                durationValue = Duration.ofSeconds(longValue);
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                durationValue = parseDurationValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                final boolean booleanCellValue = cell.getBooleanCellValue();
                // Arbitrarily assume the value is given in (seconds).
                durationValue = Duration.ofSeconds(BooleanUtils.toInteger(booleanCellValue));
                break;

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return durationValue;
    }

    /**
     * Parses a string token into a {@link Duration}. When the column's import pattern is set it
     * is tried first, falling back to ISO-8601 {@link Duration#parse} on mismatch. The value is
     * trimmed when {@code importTrim} is set; a blank value yields {@code null}.
     *
     * @param s          the raw string token
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link Duration}, or {@code null} when blank
     * @throws PxlCellCodecException if the value is not a valid duration
     */
    static Duration parseDurationValue(final String s,
                                       final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        final PxlTemporalAmountSupport.CompiledTemporalPattern importTemporalPattern = columnMeta.getImportTemporalPatternCache();
        if (Objects.nonNull(importTemporalPattern)) {
            try {
                return PxlTemporalAmountSupport.parseDurationByPattern(stringValue, importTemporalPattern);
            } catch (IllegalArgumentException ignored) {
                // On pattern mismatch, fall back to ISO-8601 parsing
            }
        }

        try {
            return Duration.parse(stringValue);
        } catch (DateTimeParseException dateTimeParseException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "Duration"), dateTimeParseException);
        }
    }

    /**
     * Writes the given value as a {@link Duration} cell and returns the exported string. A {@link String}
     * source is parsed with ISO-8601 {@link Duration#parse}; a {@link Duration} source
     * is used directly. A {@code null} result blanks the cell; otherwise the value is formatted with the
     * column's export pattern when set, else its ISO-8601 form.
     *
     * @param cell       the target cell, or {@code null} to only compute the string
     * @param object     the source value (a {@link String} or {@link Duration})
     * @param columnMeta the resolved export metadata for this column
     * @return the exported string, or {@code null} when blank
     * @throws PxlCellCodecException if the source is unsupported, the string is invalid, or the value is too large
     */
    static String buildDurationCell(final Cell cell,
                                    final Object object,
                                    final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Duration durationValue;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                durationValue = null;
            } else {
                try {
                    durationValue = Duration.parse(stringValue);
                } catch (DateTimeParseException dateTimeParseException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(stringValue), "Duration"), dateTimeParseException);
                }
            }
        } else if (object instanceof Duration) {
            durationValue = (Duration) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "Duration"));
        }

        if (Objects.isNull(durationValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else {
            final String cellString = makeDurationExportString(durationValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(cellString));
            return cellString;
        }
    }

    /**
     * Renders the export string for a {@link Duration}: when an export pattern is set the value's milliseconds are formatted
     * with {@link DurationFormatUtils}, otherwise its {@link Duration#toString()} ISO-8601 form is used; then applies
     * string-level export processing via {@link PxlStringCodec#makeExportString}.
     *
     * @param durationValue the value to render
     * @param columnMeta    resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the duration is too large to convert to milliseconds or the export pattern cannot be applied
     */
    private static String makeDurationExportString(final Duration durationValue,
                                                   final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(durationValue)) {
            return null;
        }

        final String exportPattern = columnMeta.getExportPattern();

        try {
            final long milliSeconds = durationValue.toMillis();
            String stringValue;

            if (StringUtils.isNotBlank(exportPattern)) {
                stringValue = DurationFormatUtils.formatDuration(milliSeconds, exportPattern);
            } else {
                // stringValue = DurationFormatUtils.formatDurationISO(milliSeconds);
                stringValue = durationValue.toString();
            }

            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } catch (ArithmeticException arithmeticException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_VALUE_TOO_LARGE, String.valueOf(durationValue), "Duration"), arithmeticException);
        } catch (IllegalArgumentException illegalArgumentException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(durationValue), "Duration"), illegalArgumentException);
        }
    }

}
