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
import java.text.ParseException;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Pattern;

/**
 * Codec for primitive {@code short} column values — parses cells and strings into {@code short} on import and writes
 * {@code short} into cells on export. Numeric input is range-checked against the {@code short} range (throwing on overflow)
 * and truncated to its integer part; boolean cells map to 1/0. Because {@code short} cannot be {@code null}, blank input
 * parses to {@code 0}.
 */
final class PxlPrimitiveShortCodec {

    /**
     * Prevents instantiation.
     */
    private PxlPrimitiveShortCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses an Excel cell into a {@code short}. NUMERIC cells are range-checked against the {@code short} range and
     * truncated to their integer part; STRING cells are delegated to the string overload; BOOLEAN cells map to 1 (true) or
     * 0 (false); BLANK cells yield {@code 0}.
     *
     * @param cell       the source cell
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@code short} (0 for a blank cell)
     * @throws PxlCellCodecException if the numeric value is outside the {@code short} range or the cell type is unsupported
     */
    static short parsePrimitiveShortValue(final Cell cell,
                                          final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        short shortValue = 0;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                shortValue = PxlNumberSupport.requireWithinRange(numericValue, Short.MIN_VALUE, Short.MAX_VALUE, "short").shortValue();
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                shortValue = parsePrimitiveShortValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                final boolean booleanCellValue = cell.getBooleanCellValue();
                shortValue = (short) BooleanUtils.toInteger(booleanCellValue);
                break;

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return shortValue;
    }

    /**
     * Parses a string into a {@code short}. Trims first when {@code importTrim} is enabled and returns {@code 0} for blank
     * input. When an import {@link DecimalFormat} is configured the parsed number is range-checked against the {@code short}
     * range and truncated; otherwise {@link Short#parseShort(String)} is used.
     *
     * @param s          the source string
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@code short} (0 for blank input)
     * @throws PxlCellCodecException if the string is not a valid {@code short} or is outside the {@code short} range
     */
    static short parsePrimitiveShortValue(final String s,
                                          final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return (short) 0;
        }

        short shortValue;

        final DecimalFormat importDecimalFormatter = columnMeta.getImportDecimalFormatterCache();
        if (Objects.nonNull(importDecimalFormatter)) {
            try {
                shortValue = PxlNumberSupport.requireWithinRange(importDecimalFormatter.parse(stringValue), Short.MIN_VALUE, Short.MAX_VALUE, "short").shortValue();
            } catch (ParseException parseException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "short"), parseException);
            }
        } else {
            try {
                shortValue = Short.parseShort(stringValue);
            } catch (NumberFormatException numberFormatException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "short"), numberFormatException);
            }
        }

        return shortValue;
    }

    /**
     * Writes a {@code short} value into a cell. Accepts a {@code Short} directly or a {@code String} (parsed via
     * {@link Short#parseShort(String)}; blank becomes {@code null}). A {@code null} value blanks the cell. When the column
     * is exported as text the value is formatted via {@link #makePrimitiveShortExportString} and written quote-prefixed;
     * otherwise it is written as a numeric cell.
     *
     * @param cell       the target cell (may be {@code null}, in which case only the return string is produced)
     * @param object     the source value ({@code Short} or {@code String})
     * @param columnMeta resolved export metadata for the column
     * @return the string representation of the written value, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if a string value is malformed or the object type cannot be converted to {@code short}
     */
    static String buildPrimitiveShortCell(final Cell cell,
                                          final Object object,
                                          final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Short shortValue;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                shortValue = null;
            } else {
                try {
                    shortValue = Short.parseShort(stringValue);
                } catch (NumberFormatException numberFormatException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "short"), numberFormatException);
                }
            }
        } else if (object instanceof Short) {
            shortValue = (Short) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "short"));
        }

        if (Objects.isNull(shortValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else if (columnMeta.isExportedToString()) {
            final String cellString = makePrimitiveShortExportString(shortValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> columnMeta.setQuotePrefixedCellValue(c, cellString));
            return cellString;
        } else {
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(shortValue));
            return String.valueOf(shortValue);
        }
    }

    /**
     * Renders the export string for a {@code short}: applies the export {@link DecimalFormat} when configured, otherwise
     * applies masking when an export masking pattern is set, otherwise returns the plain {@link String#valueOf(int)} form.
     *
     * @param shortValue the value to render
     * @param columnMeta resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the export pattern cannot be applied to the value
     */
    private static String makePrimitiveShortExportString(final Short shortValue,
                                                         final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(shortValue)) {
            return null;
        }

        final DecimalFormat exportDecimalFormatter = columnMeta.getExportDecimalFormatterCache();
        final Pattern exportMaskingPattern = columnMeta.getExportMaskingPattern();

        if (Objects.nonNull(exportDecimalFormatter)) {
            try {
                final String stringValue = exportDecimalFormatter.format(shortValue);
                return PxlStringCodec.makeExportString(stringValue, columnMeta);
            } catch (IllegalArgumentException illegalArgumentException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PATTERN_APPLY_FAILED, String.valueOf(shortValue)), illegalArgumentException);
            }
        } else if (Objects.nonNull(exportMaskingPattern)) {
            final String stringValue = String.valueOf(shortValue);
            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } else {
            return String.valueOf(shortValue);
        }
    }

}
