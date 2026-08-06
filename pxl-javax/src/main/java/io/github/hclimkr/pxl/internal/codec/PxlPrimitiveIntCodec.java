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
 * Codec for primitive {@code int} column values — parses cells and strings into {@code int} on import and writes {@code int}
 * into cells on export. Numeric input is range-checked against the {@code int} range (throwing on overflow) and truncated to
 * its integer part; boolean cells map to 1/0. Because {@code int} cannot be {@code null}, blank input parses to {@code 0}.
 */
final class PxlPrimitiveIntCodec {

    /**
     * Prevents instantiation.
     */
    private PxlPrimitiveIntCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses an Excel cell into an {@code int}. NUMERIC cells are range-checked against the {@code int} range and truncated
     * to their integer part; STRING cells are delegated to the string overload; BOOLEAN cells map to 1 (true) or 0 (false);
     * BLANK cells yield {@code 0}.
     *
     * @param cell       the source cell
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@code int} (0 for a blank cell)
     * @throws PxlCellCodecException if the numeric value is outside the {@code int} range or the cell type is unsupported
     */
    static int parsePrimitiveIntValue(final Cell cell,
                                      final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        int intValue = 0;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                intValue = PxlNumberSupport.requireWithinRange(numericValue, Integer.MIN_VALUE, Integer.MAX_VALUE, "int").intValue();
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                intValue = parsePrimitiveIntValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                final boolean booleanCellValue = cell.getBooleanCellValue();
                intValue = BooleanUtils.toInteger(booleanCellValue);
                break;

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return intValue;
    }

    /**
     * Parses a string into an {@code int}. Trims first when {@code importTrim} is enabled and returns {@code 0} for blank
     * input. When an import {@link DecimalFormat} is configured the parsed number is range-checked against the {@code int}
     * range and truncated; otherwise {@link Integer#parseInt(String)} is used.
     *
     * @param s          the source string
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@code int} (0 for blank input)
     * @throws PxlCellCodecException if the string is not a valid {@code int} or is outside the {@code int} range
     */
    static int parsePrimitiveIntValue(final String s,
                                      final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return (int) 0;
        }

        int intValue;

        final DecimalFormat importDecimalFormatter = columnMeta.getImportDecimalFormatterCache();
        if (Objects.nonNull(importDecimalFormatter)) {
            try {
                intValue = PxlNumberSupport.requireWithinRange(importDecimalFormatter.parse(stringValue), Integer.MIN_VALUE, Integer.MAX_VALUE, "int").intValue();
            } catch (ParseException parseException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "int"), parseException);
            }
        } else {
            try {
                intValue = Integer.parseInt(stringValue);
            } catch (NumberFormatException numberFormatException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "int"), numberFormatException);
            }
        }

        return intValue;
    }

    /**
     * Writes an {@code int} value into a cell. Accepts an {@link Integer} directly or a {@link String} (parsed via
     * {@link Integer#parseInt(String)}; blank becomes {@code null}). A {@code null} value blanks the cell. When the column
     * is exported as text the value is formatted via {@link #makePrimitiveIntExportString} and written quote-prefixed;
     * otherwise it is written as a numeric cell.
     *
     * @param cell       the target cell (may be {@code null}, in which case only the return string is produced)
     * @param object     the source value ({@link Integer} or {@link String})
     * @param columnMeta resolved export metadata for the column
     * @return the string representation of the written value, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if a string value is malformed or the object type cannot be converted to {@code int}
     */
    static String buildPrimitiveIntCell(final Cell cell,
                                        final Object object,
                                        final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Integer intValue;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                intValue = null;
            } else {
                try {
                    intValue = Integer.parseInt(stringValue);
                } catch (NumberFormatException numberFormatException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(stringValue), "int"), numberFormatException);
                }
            }
        } else if (object instanceof Integer) {
            intValue = (Integer) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "int"));
        }

        if (Objects.isNull(intValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else if (columnMeta.isExportedToString()) {
            final String cellString = makePrimitiveIntExportString(intValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> columnMeta.setQuotePrefixedCellValue(c, cellString));
            return cellString;
        } else {
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(intValue));
            return String.valueOf(intValue);
        }
    }

    /**
     * Renders the export string for an {@code int}: applies the export {@link DecimalFormat} when configured, otherwise
     * applies masking when an export masking pattern is set, otherwise returns the plain {@link String#valueOf(int)} form.
     *
     * @param intValue   the value to render
     * @param columnMeta resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the export pattern cannot be applied to the value
     */
    private static String makePrimitiveIntExportString(final Integer intValue,
                                                       final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(intValue)) {
            return null;
        }

        final DecimalFormat exportDecimalFormatter = columnMeta.getExportDecimalFormatterCache();
        final Pattern exportMaskingPattern = columnMeta.getExportMaskingPattern();

        if (Objects.nonNull(exportDecimalFormatter)) {
            try {
                final String stringValue = exportDecimalFormatter.format(intValue);
                return PxlStringCodec.makeExportString(stringValue, columnMeta);
            } catch (IllegalArgumentException illegalArgumentException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PATTERN_APPLY_FAILED, String.valueOf(intValue)), illegalArgumentException);
            }
        } else if (Objects.nonNull(exportMaskingPattern)) {
            final String stringValue = String.valueOf(intValue);
            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } else {
            return String.valueOf(intValue);
        }
    }

}
