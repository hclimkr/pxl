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
import org.apache.poi.ss.util.NumberToTextConverter;

import java.text.DecimalFormat;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Pattern;

/**
 * Codec for primitive {@code double} column values - parses cells and strings into {@code double} on import and writes
 * {@code double} into cells on export. Numeric cells are taken as-is (no range check); boolean cells map to 1.0/0.0.
 * Because {@code double} cannot be {@code null}, blank input parses to {@code 0.0}. Both import and export reject NaN and
 * Infinity, keeping the two directions symmetric, and export renders plain numeric text via {@link NumberToTextConverter}
 * to avoid scientific-notation noise.
 */
final class PxlPrimitiveDoubleCodec {

    /**
     * Prevents instantiation.
     */
    private PxlPrimitiveDoubleCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses an Excel cell into a {@code double}. NUMERIC cells are taken directly; STRING cells are delegated to the string
     * overload; BOOLEAN cells map to 1.0 (true) or 0.0 (false); BLANK cells yield {@code 0.0}.
     *
     * @param cell       the source cell
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@code double} (0.0 for a blank cell)
     * @throws PxlCellCodecException if the cell type is unsupported
     */
    static double parsePrimitiveDoubleValue(final Cell cell,
                                            final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        double doubleValue = 0.;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                doubleValue = numericValue;
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                doubleValue = parsePrimitiveDoubleValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                final boolean booleanCellValue = cell.getBooleanCellValue();
                doubleValue = (double) BooleanUtils.toInteger(booleanCellValue);
                break;

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        PxlNumberSupport.requireFiniteForImport(doubleValue, "double");

        return doubleValue;
    }

    /**
     * Parses a string into a {@code double}. Trims first when {@code importTrim} is enabled and returns {@code 0.0} for blank
     * input. When an import {@link DecimalFormat} is configured the whole string must match the pattern
     * ({@code PxlNumberSupport.parseFullyAsNumber}) and its parsed value is used; otherwise {@link Double#parseDouble(String)}
     * is used.
     *
     * @param s          the source string
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@code double} (0.0 for blank input)
     * @throws PxlCellCodecException if the string is not a valid {@code double}
     */
    static double parsePrimitiveDoubleValue(final String s,
                                            final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return (double) 0.;
        }

        double doubleValue;

        final DecimalFormat importDecimalFormatter = columnMeta.getImportDecimalFormatterCache();
        if (Objects.nonNull(importDecimalFormatter)) {
            doubleValue = PxlNumberSupport.parseFullyAsNumber(importDecimalFormatter, stringValue, "double").doubleValue();
        } else {
            try {
                doubleValue = Double.parseDouble(stringValue);
            } catch (NumberFormatException numberFormatException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "double"), numberFormatException);
            }
        }

        PxlNumberSupport.requireFiniteForImport(doubleValue, "double");

        return doubleValue;
    }

    /**
     * Writes a {@code double} value into a cell. Accepts a {@link Double} directly or a {@link String} (parsed via
     * {@link Double#parseDouble(String)}; blank becomes {@code null}). A {@code null} value blanks the cell; NaN or Infinity
     * is rejected. When the column is exported as text the value is formatted via {@link #makePrimitiveDoubleExportString}
     * and written quote-prefixed; otherwise it is written as a numeric cell.
     *
     * @param cell       the target cell (may be {@code null}, in which case only the return string is produced)
     * @param object     the source value ({@link Double} or {@link String})
     * @param columnMeta resolved export metadata for the column
     * @return the string representation of the written value, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if a string value is malformed, the object type cannot be converted to {@code double},
     *                               or the value is NaN or Infinity
     */
    static String buildPrimitiveDoubleCell(final Cell cell,
                                           final Object object,
                                           final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Double doubleValue;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                doubleValue = null;
            } else {
                try {
                    doubleValue = Double.parseDouble(stringValue);
                } catch (NumberFormatException numberFormatException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(stringValue), "double"), numberFormatException);
                }
            }
        } else if (object instanceof Double) {
            doubleValue = (Double) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "double"));
        }

        if (Objects.isNull(doubleValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        }

        PxlNumberSupport.requireFiniteForExport(doubleValue);

        if (columnMeta.isExportedToString()) {
            final String cellString = makePrimitiveDoubleExportString(doubleValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> columnMeta.setQuotePrefixedCellValue(c, cellString));
            return cellString;
        } else {
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(doubleValue));
            return NumberToTextConverter.toText(doubleValue);
        }
    }

    /**
     * Renders the export string for a {@code double}: applies the export {@link DecimalFormat} when configured, otherwise
     * applies masking over the {@link NumberToTextConverter} text when an export masking pattern is set, otherwise returns
     * that plain numeric text.
     *
     * @param doubleValue the value to render
     * @param columnMeta  resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the export pattern cannot be applied to the value
     */
    private static String makePrimitiveDoubleExportString(final Double doubleValue,
                                                          final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(doubleValue)) {
            return null;
        }

        final DecimalFormat exportDecimalFormatter = columnMeta.getExportDecimalFormatterCache();
        final Pattern exportMaskingPattern = columnMeta.getExportMaskingPattern();

        if (Objects.nonNull(exportDecimalFormatter)) {
            try {
                final String stringValue = exportDecimalFormatter.format(doubleValue);
                return PxlStringCodec.makeExportString(stringValue, columnMeta);
            } catch (IllegalArgumentException illegalArgumentException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PATTERN_APPLY_FAILED, String.valueOf(doubleValue)), illegalArgumentException);
            }
        } else if (Objects.nonNull(exportMaskingPattern)) {
            final String stringValue = NumberToTextConverter.toText(doubleValue);
            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } else {
            return NumberToTextConverter.toText(doubleValue);
        }
    }

}
