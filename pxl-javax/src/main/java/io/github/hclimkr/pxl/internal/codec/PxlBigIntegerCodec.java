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

import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Pattern;

/**
 * Codec for {@link BigInteger} column values — parses cells and strings into {@link BigInteger} on import and writes
 * {@link BigInteger} into cells on export. A numeric cell is converted via {@code BigDecimal.valueOf(double)} and truncated
 * to its integer part (so its precision is limited by the underlying {@code double}); string input is parsed exactly via
 * {@code new BigInteger(String)}. Boolean cells map to 1/0. Export always writes the value as text
 * ({@link BigInteger#toString()}, quote-prefixed) to preserve full precision rather than as a lossy numeric cell.
 */
final class PxlBigIntegerCodec {

    /**
     * Prevents instantiation.
     */
    private PxlBigIntegerCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses an Excel cell into a {@link BigInteger}. NUMERIC cells are converted via {@code BigDecimal.valueOf(double)} and
     * truncated to their integer part; STRING cells are delegated to the string overload; BOOLEAN cells map to 1 (true) or
     * 0 (false); BLANK cells yield {@code null}.
     *
     * @param cell       the source cell
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@link BigInteger}, or {@code null} for a blank cell
     * @throws PxlCellCodecException if the cell type is unsupported
     */
    static BigInteger parseBigIntegerValue(final Cell cell,
                                           final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        BigInteger bigIntegerValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                bigIntegerValue = PxlNumberSupport.requireFinite(numericValue, "BigInteger").toBigInteger();
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                bigIntegerValue = parseBigIntegerValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                final boolean booleanCellValue = cell.getBooleanCellValue();
                bigIntegerValue = BigDecimal.valueOf(BooleanUtils.toInteger(booleanCellValue)).toBigInteger();
                break;

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return bigIntegerValue;
    }

    /**
     * Parses a string into a {@link BigInteger}. Trims first when {@code importTrim} is enabled and returns {@code null} for
     * blank input. When an import {@link DecimalFormat} is configured (parsing to {@link BigDecimal}) its result is truncated
     * to a {@link BigInteger}; otherwise {@code new BigInteger(String)} parses the value exactly.
     *
     * @param s          the source string
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@link BigInteger}, or {@code null} for blank input
     * @throws PxlCellCodecException if the string is not a valid {@link BigInteger}
     */
    static BigInteger parseBigIntegerValue(final String s,
                                           final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        BigInteger bigIntegerValue;

        final DecimalFormat importDecimalFormatter = columnMeta.getImportDecimalFormatterCache();
        if (Objects.nonNull(importDecimalFormatter)) {
            try {
                bigIntegerValue = ((BigDecimal) importDecimalFormatter.parse(stringValue)).toBigInteger();
            } catch (ParseException parseException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "BigInteger"), parseException);
            }
        } else {
            try {
                bigIntegerValue = new BigInteger(stringValue);
            } catch (NumberFormatException numberFormatException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "BigInteger"), numberFormatException);
            }
        }

        return bigIntegerValue;
    }

    /**
     * Writes a {@link BigInteger} value into a cell. Accepts a {@link BigInteger} directly or a {@link String} (parsed via
     * {@code new BigInteger(String)}; blank becomes {@code null}). A {@code null} value blanks the cell. When the column is
     * exported as text the value is formatted via {@link #makeBigIntegerExportString}; otherwise its
     * {@link BigInteger#toString()} form is written quote-prefixed (as text, to preserve precision).
     *
     * @param cell       the target cell (may be {@code null}, in which case only the return string is produced)
     * @param object     the source value ({@link BigInteger} or {@link String})
     * @param columnMeta resolved export metadata for the column
     * @return the string representation of the written value, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if a string value is malformed or the object type cannot be converted to {@link BigInteger}
     */
    static String buildBigIntegerCell(final Cell cell,
                                      final Object object,
                                      final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        BigInteger bigIntegerValue;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                bigIntegerValue = null;
            } else {
                try {
                    // bigIntegerValue = BigDecimal.valueOf(Double.parseDouble(stringValue)).toBigInteger();
                    bigIntegerValue = new BigInteger(stringValue);
                } catch (NumberFormatException numberFormatException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(stringValue), "BigInteger"), numberFormatException);
                }
            }
        } else if (object instanceof BigInteger) {
            bigIntegerValue = (BigInteger) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "BigInteger"));
        }

        if (Objects.isNull(bigIntegerValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else if (columnMeta.isExportedToString()) {
            final String cellString = makeBigIntegerExportString(bigIntegerValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> columnMeta.setQuotePrefixedCellValue(c, cellString));
            return cellString;
        } else {
            // Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(bigIntegerValue.doubleValue()));
            // return NumberToTextConverter.toText(bigIntegerValue.doubleValue());
            final String cellString = bigIntegerValue.toString();
            Optional.ofNullable(cell).ifPresent(c -> columnMeta.setQuotePrefixedCellValue(c, cellString));
            return cellString;
        }
    }

    /**
     * Renders the export string for a {@link BigInteger}: applies the export {@link DecimalFormat} when configured, otherwise
     * applies masking over the {@link BigInteger#toString()} form when an export masking pattern is set, otherwise returns
     * that string form.
     *
     * @param bigIntegerValue the value to render
     * @param columnMeta      resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the export pattern cannot be applied to the value
     */
    private static String makeBigIntegerExportString(final BigInteger bigIntegerValue,
                                                     final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(bigIntegerValue)) {
            return null;
        }

        final DecimalFormat exportDecimalFormatter = columnMeta.getExportDecimalFormatterCache();
        final Pattern exportMaskingPattern = columnMeta.getExportMaskingPattern();

        if (Objects.nonNull(exportDecimalFormatter)) {
            try {
                final String stringValue = exportDecimalFormatter.format(bigIntegerValue);
                return PxlStringCodec.makeExportString(stringValue, columnMeta);
            } catch (IllegalArgumentException illegalArgumentException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PATTERN_APPLY_FAILED, String.valueOf(bigIntegerValue)), illegalArgumentException);
            }
        } else if (Objects.nonNull(exportMaskingPattern)) {
            // final String stringValue = NumberToTextConverter.toText(bigIntegerValue.doubleValue());
            final String stringValue = bigIntegerValue.toString();
            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } else {
            // return NumberToTextConverter.toText(bigIntegerValue.doubleValue());
            return bigIntegerValue.toString();
        }
    }

}
