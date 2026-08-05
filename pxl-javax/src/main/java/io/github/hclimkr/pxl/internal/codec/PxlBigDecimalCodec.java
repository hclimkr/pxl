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
import java.text.DecimalFormat;
import java.text.ParseException;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Pattern;

/**
 * Codec for {@link BigDecimal} column values — parses cells and strings into {@link BigDecimal} on import and writes
 * {@link BigDecimal} into cells on export. A numeric cell is converted via {@code BigDecimal.valueOf(double)} (so its
 * precision is limited by the underlying {@code double}); string input is parsed exactly via {@code new BigDecimal(String)}.
 * Boolean cells map to 1/0. Export always writes the value as text ({@link BigDecimal#toPlainString()}, quote-prefixed) to
 * preserve full precision rather than as a lossy numeric cell.
 */
final class PxlBigDecimalCodec {

    /**
     * Prevents instantiation.
     */
    private PxlBigDecimalCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses an Excel cell into a {@link BigDecimal}. NUMERIC cells are converted via {@code BigDecimal.valueOf(double)};
     * STRING cells are delegated to the string overload; BOOLEAN cells map to 1 (true) or 0 (false); BLANK cells yield {@code null}.
     *
     * @param cell       the source cell
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@link BigDecimal}, or {@code null} for a blank cell
     * @throws PxlCellCodecException if the cell type is unsupported
     */
    static BigDecimal parseBigDecimalValue(final Cell cell,
                                           final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        BigDecimal bigDecimalValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                bigDecimalValue = PxlNumberSupport.requireFinite(numericValue, "BigDecimal");
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                bigDecimalValue = parseBigDecimalValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                final boolean booleanCellValue = cell.getBooleanCellValue();
                bigDecimalValue = BigDecimal.valueOf(BooleanUtils.toInteger(booleanCellValue));
                break;

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return bigDecimalValue;
    }

    /**
     * Parses a string into a {@link BigDecimal}. Trims first when {@code importTrim} is enabled and returns {@code null} for
     * blank input. When an import {@link DecimalFormat} is configured (parsing to {@link BigDecimal}) its result is used;
     * otherwise {@code new BigDecimal(String)} parses the value exactly.
     *
     * @param s          the source string
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@link BigDecimal}, or {@code null} for blank input
     * @throws PxlCellCodecException if the string is not a valid {@link BigDecimal}
     */
    static BigDecimal parseBigDecimalValue(final String s,
                                           final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        BigDecimal bigDecimalValue;

        final DecimalFormat importDecimalFormatter = columnMeta.getImportDecimalFormatterCache();
        if (Objects.nonNull(importDecimalFormatter)) {
            try {
                bigDecimalValue = (BigDecimal) importDecimalFormatter.parse(stringValue);
            } catch (ParseException parseException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "BigDecimal"), parseException);
            }
        } else {
            try {
                bigDecimalValue = new BigDecimal(stringValue);
            } catch (NumberFormatException numberFormatException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "BigDecimal"), numberFormatException);
            }
        }

        return bigDecimalValue;
    }

    /**
     * Writes a {@link BigDecimal} value into a cell. Accepts a {@link BigDecimal} directly or a {@link String} (parsed via
     * {@code new BigDecimal(String)}; blank becomes {@code null}). A {@code null} value blanks the cell. When the column is
     * exported as text the value is formatted via {@link #makeBigDecimalExportString}; otherwise its
     * {@link BigDecimal#toPlainString()} form is written quote-prefixed (as text, to preserve precision).
     *
     * @param cell       the target cell (may be {@code null}, in which case only the return string is produced)
     * @param object     the source value ({@link BigDecimal} or {@link String})
     * @param columnMeta resolved export metadata for the column
     * @return the string representation of the written value, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if a string value is malformed or the object type cannot be converted to {@link BigDecimal}
     */
    static String buildBigDecimalCell(final Cell cell,
                                      final Object object,
                                      final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        BigDecimal bigDecimalValue;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                bigDecimalValue = null;
            } else {
                try {
                    // bigDecimalValue = BigDecimal.valueOf(Double.parseDouble(stringValue));
                    bigDecimalValue = new BigDecimal(stringValue);
                } catch (NumberFormatException numberFormatException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "BigDecimal"), numberFormatException);
                }
            }
        } else if (object instanceof BigDecimal) {
            bigDecimalValue = (BigDecimal) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "BigDecimal"));
        }

        if (Objects.isNull(bigDecimalValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else if (columnMeta.isExportedToString()) {
            final String cellString = makeBigDecimalExportString(bigDecimalValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> columnMeta.setQuotePrefixedCellValue(c, cellString));
            return cellString;
        } else {
            // Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(bigDecimalValue.doubleValue()));
            // return NumberToTextConverter.toText(bigDecimalValue.doubleValue());
            final String cellString = bigDecimalValue.toPlainString();
            Optional.ofNullable(cell).ifPresent(c -> columnMeta.setQuotePrefixedCellValue(c, cellString));
            return cellString;
        }
    }

    /**
     * Renders the export string for a {@link BigDecimal}: applies the export {@link DecimalFormat} when configured, otherwise
     * applies masking over the {@link BigDecimal#toPlainString()} form when an export masking pattern is set, otherwise
     * returns that plain-string form.
     *
     * @param bigDecimalValue the value to render
     * @param columnMeta      resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the export pattern cannot be applied to the value
     */
    private static String makeBigDecimalExportString(final BigDecimal bigDecimalValue,
                                                     final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(bigDecimalValue)) {
            return null;
        }

        final DecimalFormat exportDecimalFormatter = columnMeta.getExportDecimalFormatterCache();
        final Pattern exportMaskingPattern = columnMeta.getExportMaskingPattern();

        if (Objects.nonNull(exportDecimalFormatter)) {
            try {
                final String stringValue = exportDecimalFormatter.format(bigDecimalValue);
                return PxlStringCodec.makeExportString(stringValue, columnMeta);
            } catch (IllegalArgumentException illegalArgumentException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PATTERN_APPLY_FAILED, String.valueOf(bigDecimalValue)), illegalArgumentException);
            }
        } else if (Objects.nonNull(exportMaskingPattern)) {
            // final String stringValue = NumberToTextConverter.toText(bigDecimalValue.doubleValue());
            final String stringValue = bigDecimalValue.toPlainString();
            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } else {
            // return NumberToTextConverter.toText(bigDecimalValue.doubleValue());
            return bigDecimalValue.toPlainString();
        }
    }

}
