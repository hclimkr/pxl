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
 * Codec for primitive {@code byte} column values — parses cells and strings into {@code byte} on import and writes {@code byte}
 * into cells on export. Numeric input is range-checked against the {@code byte} range (throwing on overflow) and truncated to
 * its integer part; boolean cells map to 1/0. Because {@code byte} cannot be {@code null}, blank input parses to {@code 0}.
 */
final class PxlPrimitiveByteCodec {

    /**
     * Prevents instantiation.
     */
    private PxlPrimitiveByteCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses an Excel cell into a {@code byte}. NUMERIC cells are range-checked against the {@code byte} range and truncated
     * to their integer part; STRING cells are delegated to the string overload; BOOLEAN cells map to 1 (true) or 0 (false);
     * BLANK cells yield {@code 0}.
     *
     * @param cell       the source cell
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@code byte} (0 for a blank cell)
     * @throws PxlCellCodecException if the numeric value is outside the {@code byte} range or the cell type is unsupported
     */
    static byte parsePrimitiveByteValue(final Cell cell,
                                        final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        byte byteValue = 0;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                byteValue = PxlNumberSupport.requireWithinRange(numericValue, Byte.MIN_VALUE, Byte.MAX_VALUE, "byte").byteValue();
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                byteValue = parsePrimitiveByteValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                final boolean booleanCellValue = cell.getBooleanCellValue();
                byteValue = (byte) BooleanUtils.toInteger(booleanCellValue);
                break;

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return byteValue;
    }

    /**
     * Parses a string into a {@code byte}. Trims first when {@code importTrim} is enabled and returns {@code 0} for blank
     * input. When an import {@link DecimalFormat} is configured the parsed number is range-checked against the {@code byte}
     * range and truncated; otherwise {@link Byte#parseByte(String)} is used.
     *
     * @param s          the source string
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@code byte} (0 for blank input)
     * @throws PxlCellCodecException if the string is not a valid {@code byte} or is outside the {@code byte} range
     */
    static byte parsePrimitiveByteValue(final String s,
                                        final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return (byte) 0;
        }

        byte byteValue;

        final DecimalFormat importDecimalFormatter = columnMeta.getImportDecimalFormatterCache();
        if (Objects.nonNull(importDecimalFormatter)) {
            try {
                byteValue = PxlNumberSupport.requireWithinRange(importDecimalFormatter.parse(stringValue), Byte.MIN_VALUE, Byte.MAX_VALUE, "byte").byteValue();
            } catch (ParseException parseException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "byte"), parseException);
            }
        } else {
            try {
                byteValue = Byte.parseByte(stringValue);
            } catch (NumberFormatException numberFormatException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "byte"), numberFormatException);
            }
        }

        return byteValue;
    }

    /**
     * Writes a {@code byte} value into a cell. Accepts a {@link Byte} directly or a {@link String} (parsed via
     * {@link Byte#parseByte(String)}; blank becomes {@code null}). A {@code null} value blanks the cell. When the column
     * is exported as text the value is formatted via {@link #makePrimitiveByteExportString} and written quote-prefixed;
     * otherwise it is written as a numeric cell.
     *
     * @param cell       the target cell (may be {@code null}, in which case only the return string is produced)
     * @param object     the source value ({@link Byte} or {@link String})
     * @param columnMeta resolved export metadata for the column
     * @return the string representation of the written value, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if a string value is malformed or the object type cannot be converted to {@code byte}
     */
    static String buildPrimitiveByteCell(final Cell cell,
                                         final Object object,
                                         final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Byte byteValue;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                byteValue = null;
            } else {
                try {
                    byteValue = Byte.parseByte(stringValue);
                } catch (NumberFormatException numberFormatException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(stringValue), "byte"), numberFormatException);
                }
            }
        } else if (object instanceof Byte) {
            byteValue = (Byte) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "byte"));
        }

        if (Objects.isNull(byteValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else if (columnMeta.isExportedToString()) {
            final String cellString = makePrimitiveByteExportString(byteValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> columnMeta.setQuotePrefixedCellValue(c, cellString));
            return cellString;
        } else {
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(byteValue));
            return String.valueOf(byteValue);
        }
    }

    /**
     * Renders the export string for a {@code byte}: applies the export {@link DecimalFormat} when configured, otherwise
     * applies masking when an export masking pattern is set, otherwise returns the plain {@link String#valueOf(int)} form.
     *
     * @param byteValue  the value to render
     * @param columnMeta resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the export pattern cannot be applied to the value
     */
    private static String makePrimitiveByteExportString(final Byte byteValue,
                                                        final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(byteValue)) {
            return null;
        }

        final DecimalFormat exportDecimalFormatter = columnMeta.getExportDecimalFormatterCache();
        final Pattern exportMaskingPattern = columnMeta.getExportMaskingPattern();

        if (Objects.nonNull(exportDecimalFormatter)) {
            try {
                final String stringValue = exportDecimalFormatter.format(byteValue);
                return PxlStringCodec.makeExportString(stringValue, columnMeta);
            } catch (IllegalArgumentException illegalArgumentException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PATTERN_APPLY_FAILED, String.valueOf(byteValue)), illegalArgumentException);
            }
        } else if (Objects.nonNull(exportMaskingPattern)) {
            final String stringValue = String.valueOf(byteValue);
            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } else {
            return String.valueOf(byteValue);
        }
    }

}
