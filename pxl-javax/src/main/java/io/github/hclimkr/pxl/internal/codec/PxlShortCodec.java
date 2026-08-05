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
 * Codec for {@link Short} column values — parses cells and strings into {@link Short} on import and writes {@link Short}
 * into cells on export. Numeric input is range-checked against the {@link Short} range (throwing on overflow) and truncated
 * to its integer part; boolean cells map to 1/0.
 */
final class PxlShortCodec {

    /**
     * Prevents instantiation.
     */
    private PxlShortCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses an Excel cell into a {@link Short}. NUMERIC cells are range-checked against the {@link Short} range and
     * truncated to their integer part; STRING cells are delegated to the string overload; BOOLEAN cells map to 1 (true) or
     * 0 (false); BLANK cells yield {@code null}.
     *
     * @param cell       the source cell
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@link Short}, or {@code null} for a blank cell
     * @throws PxlCellCodecException if the numeric value is outside the {@link Short} range or the cell type is unsupported
     */
    static Short parseShortValue(final Cell cell,
                                 final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Short shortValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                shortValue = PxlNumberSupport.requireWithinRange(numericValue, Short.MIN_VALUE, Short.MAX_VALUE, "Short").shortValue();
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                shortValue = parseShortValue(stringCellValue, columnMeta);
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
     * Parses a string into a {@link Short}. Trims first when {@code importTrim} is enabled and returns {@code null} for
     * blank input. When an import {@link DecimalFormat} is configured the parsed number is range-checked against the
     * {@link Short} range and truncated; otherwise {@link Short#parseShort(String)} is used.
     *
     * @param s          the source string
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@link Short}, or {@code null} for blank input
     * @throws PxlCellCodecException if the string is not a valid {@link Short} or is outside the {@link Short} range
     */
    static Short parseShortValue(final String s,
                                 final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        Short shortValue;

        final DecimalFormat importDecimalFormatter = columnMeta.getImportDecimalFormatterCache();
        if (Objects.nonNull(importDecimalFormatter)) {
            try {
                shortValue = PxlNumberSupport.requireWithinRange(importDecimalFormatter.parse(stringValue), Short.MIN_VALUE, Short.MAX_VALUE, "Short").shortValue();
            } catch (ParseException parseException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "Short"), parseException);
            }
        } else {
            try {
                shortValue = Short.parseShort(stringValue);
            } catch (NumberFormatException numberFormatException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "Short"), numberFormatException);
            }
        }

        return shortValue;
    }

    /**
     * Writes a {@link Short} value into a cell. Accepts a {@link Short} directly or a {@link String} (parsed via
     * {@link Short#parseShort(String)}; blank becomes {@code null}). A {@code null} value blanks the cell. When the column
     * is exported as text the value is formatted via {@link #makeShortExportString} and written quote-prefixed; otherwise
     * it is written as a numeric cell.
     *
     * @param cell       the target cell (may be {@code null}, in which case only the return string is produced)
     * @param object     the source value ({@link Short} or {@link String})
     * @param columnMeta resolved export metadata for the column
     * @return the string representation of the written value, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if a string value is malformed or the object type cannot be converted to {@link Short}
     */
    static String buildShortCell(final Cell cell,
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
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "Short"), numberFormatException);
                }
            }
        } else if (object instanceof Short) {
            shortValue = (Short) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "Short"));
        }

        if (Objects.isNull(shortValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else if (columnMeta.isExportedToString()) {
            final String cellString = makeShortExportString(shortValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> columnMeta.setQuotePrefixedCellValue(c, cellString));
            return cellString;
        } else {
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(shortValue));
            return String.valueOf(shortValue);
        }
    }

    /**
     * Renders the export string for a {@link Short}: applies the export {@link DecimalFormat} when configured, otherwise
     * applies masking when an export masking pattern is set, otherwise returns the plain {@link String#valueOf(int)} form.
     *
     * @param shortValue the value to render
     * @param columnMeta resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the export pattern cannot be applied to the value
     */
    private static String makeShortExportString(final Short shortValue,
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
