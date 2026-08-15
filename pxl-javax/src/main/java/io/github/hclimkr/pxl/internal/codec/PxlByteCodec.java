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
 * Codec for {@link Byte} column values - parses cells and strings into {@link Byte} on import and writes {@link Byte}
 * into cells on export. Numeric input is range-checked against the {@link Byte} range (throwing on overflow) and truncated
 * to its integer part; boolean cells map to 1/0.
 */
final class PxlByteCodec {

    /**
     * Prevents instantiation.
     */
    private PxlByteCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses an Excel cell into a {@link Byte}. NUMERIC cells are range-checked against the {@link Byte} range and truncated
     * to their integer part; STRING cells are delegated to the string overload; BOOLEAN cells map to 1 (true) or 0 (false);
     * BLANK cells yield {@code null}.
     *
     * @param cell       the source cell
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@link Byte}, or {@code null} for a blank cell
     * @throws PxlCellCodecException if the numeric value is outside the {@link Byte} range or the cell type is unsupported
     */
    static Byte parseByteValue(final Cell cell,
                               final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Byte byteValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                byteValue = PxlNumberSupport.requireWithinRange(numericValue, Byte.MIN_VALUE, Byte.MAX_VALUE, "Byte").byteValue();
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                byteValue = parseByteValue(stringCellValue, columnMeta);
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
     * Parses a string into a {@link Byte}. Trims first when {@code importTrim} is enabled and returns {@code null} for blank
     * input. When an import {@link DecimalFormat} is configured the whole string must match the pattern
     * ({@code PxlNumberSupport.parseFullyAsNumber}) and the parsed number is range-checked against the {@link Byte} range and
     * truncated; otherwise {@link Byte#parseByte(String)} is used.
     *
     * @param s          the source string
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@link Byte}, or {@code null} for blank input
     * @throws PxlCellCodecException if the string is not a valid {@link Byte} or is outside the {@link Byte} range
     */
    static Byte parseByteValue(final String s,
                               final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        Byte byteValue;

        final DecimalFormat importDecimalFormatter = columnMeta.getImportDecimalFormatterCache();
        if (Objects.nonNull(importDecimalFormatter)) {
            byteValue = PxlNumberSupport.requireWithinRange(PxlNumberSupport.parseFullyAsNumber(importDecimalFormatter, stringValue, "Byte"), Byte.MIN_VALUE, Byte.MAX_VALUE, "Byte").byteValue();
        } else {
            try {
                byteValue = Byte.parseByte(stringValue);
            } catch (NumberFormatException numberFormatException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "Byte"), numberFormatException);
            }
        }

        return byteValue;
    }

    /**
     * Writes a {@link Byte} value into a cell. Accepts a {@link Byte} directly or a {@link String} (parsed via
     * {@link Byte#parseByte(String)}; blank becomes {@code null}). A {@code null} value blanks the cell. When the column
     * is exported as text the value is formatted via {@link #makeByteExportString} and written quote-prefixed; otherwise
     * it is written as a numeric cell.
     *
     * @param cell       the target cell (may be {@code null}, in which case only the return string is produced)
     * @param object     the source value ({@link Byte} or {@link String})
     * @param columnMeta resolved export metadata for the column
     * @return the string representation of the written value, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if a string value is malformed or the object type cannot be converted to {@link Byte}
     */
    static String buildByteCell(final Cell cell,
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
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(stringValue), "Byte"), numberFormatException);
                }
            }
        } else if (object instanceof Byte) {
            byteValue = (Byte) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "Byte"));
        }

        if (Objects.isNull(byteValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else if (columnMeta.isExportedToString()) {
            final String cellString = makeByteExportString(byteValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> columnMeta.setQuotePrefixedCellValue(c, cellString));
            return cellString;
        } else {
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(byteValue));
            return String.valueOf(byteValue);
        }
    }

    /**
     * Renders the export string for a {@link Byte}: applies the export {@link DecimalFormat} when configured, otherwise
     * applies masking when an export masking pattern is set, otherwise returns the plain {@link String#valueOf(int)} form.
     *
     * @param byteValue  the value to render
     * @param columnMeta resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the export pattern cannot be applied to the value
     */
    private static String makeByteExportString(final Byte byteValue,
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
