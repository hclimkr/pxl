package io.github.hclimkr.pxl.internal.codec;

import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlExportColumnMeta;
import io.github.hclimkr.pxl.internal.meta.PxlImportColumnMeta;
import io.github.hclimkr.pxl.internal.support.PxlNumberSupport;
import org.apache.commons.lang3.BooleanUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.text.DecimalFormat;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Pattern;

/**
 * Codec for {@link Long} column values - parses cells and strings into {@link Long} on import and writes {@link Long}
 * into cells on export. Numeric input is range-checked against the {@link Long} range (throwing on overflow) and truncated
 * to its integer part; note that a numeric cell is a {@code double}, so magnitudes beyond 2^53 lose precision. Boolean cells map to 1/0.
 */
final class PxlLongCodec {

    /**
     * Prevents instantiation.
     */
    private PxlLongCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses an Excel cell into a {@link Long}. NUMERIC cells are range-checked against the {@link Long} range and truncated
     * to their integer part (magnitudes beyond 2^53 already lose precision as a {@code double}); STRING cells are delegated
     * to the string overload; BOOLEAN cells map to 1 (true) or 0 (false); BLANK cells yield {@code null}.
     *
     * @param cell       the source cell
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@link Long}, or {@code null} for a blank cell
     * @throws PxlCellCodecException if the numeric value is outside the {@link Long} range or the cell type is unsupported
     */
    static Long parseLongValue(final Cell cell,
                               final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Long longValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                longValue = PxlNumberSupport.requireWithinRange(numericValue, Long.MIN_VALUE, Long.MAX_VALUE, "Long").longValue();
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                longValue = parseLongValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                final boolean booleanCellValue = cell.getBooleanCellValue();
                longValue = (long) BooleanUtils.toInteger(booleanCellValue);
                break;

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return longValue;
    }

    /**
     * Parses a string into a {@link Long}. Trims first when {@code importTrim} is enabled and returns {@code null} for blank
     * input. When an import {@link DecimalFormat} is configured the whole string must match the pattern
     * ({@code PxlNumberSupport.parseFullyAsNumber}) and the parsed number is range-checked against the {@link Long} range and
     * truncated; otherwise {@link Long#parseLong(String)} is used (preserving full precision).
     *
     * @param s          the source string
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@link Long}, or {@code null} for blank input
     * @throws PxlCellCodecException if the string is not a valid {@link Long} or is outside the {@link Long} range
     */
    static Long parseLongValue(final String s,
                               final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        Long longValue;

        final DecimalFormat importDecimalFormatter = columnMeta.getImportDecimalFormatterCache();
        if (Objects.nonNull(importDecimalFormatter)) {
            longValue = PxlNumberSupport.requireWithinRange(PxlNumberSupport.parseFullyAsNumber(importDecimalFormatter, stringValue, "Long"), Long.MIN_VALUE, Long.MAX_VALUE, "Long").longValue();
        } else {
            try {
                longValue = Long.parseLong(stringValue);
            } catch (NumberFormatException numberFormatException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "Long"), numberFormatException);
            }
        }

        return longValue;
    }

    /**
     * Writes a {@link Long} value into a cell. Accepts a {@link Long} directly or a {@link String} (parsed via
     * {@link Long#parseLong(String)}; blank becomes {@code null}). A {@code null} value blanks the cell. When the column
     * is exported as text the value is formatted via {@link #makeLongExportString} and written quote-prefixed; otherwise
     * it is written as a numeric cell (subject to the 2^53 {@code double} precision limit).
     *
     * @param cell       the target cell (may be {@code null}, in which case only the return string is produced)
     * @param object     the source value ({@link Long} or {@link String})
     * @param columnMeta resolved export metadata for the column
     * @return the string representation of the written value, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if a string value is malformed or the object type cannot be converted to {@link Long}
     */
    static String buildLongCell(final Cell cell,
                                final Object object,
                                final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Long longValue;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                longValue = null;
            } else {
                try {
                    longValue = Long.parseLong(stringValue);
                } catch (NumberFormatException numberFormatException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(stringValue), "Long"), numberFormatException);
                }
            }
        } else if (object instanceof Long) {
            longValue = (Long) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "Long"));
        }

        if (Objects.isNull(longValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else if (columnMeta.isExportedToString()) {
            final String cellString = makeLongExportString(longValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> columnMeta.setQuotePrefixedCellValue(c, cellString));
            return cellString;
        } else {
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(longValue));
            return String.valueOf(longValue);
        }
    }

    /**
     * Renders the export string for a {@link Long}: applies the export {@link DecimalFormat} when configured, otherwise
     * applies masking when an export masking pattern is set, otherwise returns the plain {@link String#valueOf(long)} form.
     *
     * @param longValue  the value to render
     * @param columnMeta resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the export pattern cannot be applied to the value
     */
    private static String makeLongExportString(final Long longValue,
                                               final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(longValue)) {
            return null;
        }

        final DecimalFormat exportDecimalFormatter = columnMeta.getExportDecimalFormatterCache();
        final Pattern exportMaskingPattern = columnMeta.getExportMaskingPattern();

        if (Objects.nonNull(exportDecimalFormatter)) {
            try {
                final String stringValue = exportDecimalFormatter.format(longValue);
                return PxlStringCodec.makeExportString(stringValue, columnMeta);
            } catch (IllegalArgumentException illegalArgumentException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PATTERN_APPLY_FAILED, String.valueOf(longValue)), illegalArgumentException);
            }
        } else if (Objects.nonNull(exportMaskingPattern)) {
            final String stringValue = String.valueOf(longValue);
            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } else {
            return String.valueOf(longValue);
        }
    }

}
