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
 * Codec for {@link Integer} column values - writes {@link Integer} into cells on export and parses cells and
 * strings into {@link Integer} on import. Numeric input is range-checked against the {@link Integer} range (throwing
 * on overflow) and truncated to its integer part; boolean cells map to 1/0.
 */
final class PxlIntegerCodec {

    /**
     * Prevents instantiation.
     */
    private PxlIntegerCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Writes an {@link Integer} value into a cell. Accepts an {@link Integer} directly or a {@link String} (parsed via
     * {@link Integer#parseInt(String)}; blank becomes {@code null}). A {@code null} value blanks the cell. When the column
     * is exported as text the value is formatted via {@link #makeIntegerExportString} and written quote-prefixed; otherwise
     * it is written as a numeric cell.
     *
     * @param cell       the target cell (may be {@code null}, in which case only the return string is produced)
     * @param object     the source value ({@link Integer} or {@link String})
     * @param columnMeta resolved export metadata for the column
     * @return the string representation of the written value, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if a string value is malformed or the object type cannot be converted to {@link Integer}
     */
    static String buildIntegerCell(final Cell cell,
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
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(stringValue), "Integer"), numberFormatException);
                }
            }
        } else if (object instanceof Integer) {
            intValue = (Integer) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "Integer"));
        }

        if (Objects.isNull(intValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else if (columnMeta.isExportedToString()) {
            final String cellString = makeIntegerExportString(intValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> columnMeta.setQuotePrefixedCellValue(c, cellString));
            return cellString;
        } else {
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(intValue));
            return String.valueOf(intValue);
        }
    }

    /**
     * Renders the export string for an {@link Integer}: applies the export {@link DecimalFormat} when configured, otherwise
     * applies masking when an export masking pattern is set, otherwise returns the plain {@link String#valueOf(int)} form.
     *
     * @param intValue   the value to render
     * @param columnMeta resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the export pattern cannot be applied to the value
     */
    private static String makeIntegerExportString(final Integer intValue,
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

    /**
     * Parses an Excel cell into an {@link Integer}. NUMERIC cells are range-checked against the {@link Integer} range
     * and truncated to their integer part; STRING cells are delegated to the string overload; BOOLEAN cells map to
     * 1 (true) or 0 (false); BLANK cells yield {@code null}.
     *
     * @param cell       the source cell
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@link Integer}, or {@code null} for a blank cell
     * @throws PxlCellCodecException if the numeric value is outside the {@link Integer} range or the cell type is unsupported
     */
    static Integer parseIntegerValue(final Cell cell,
                                     final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Integer integerValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                integerValue = PxlNumberSupport.requireWithinRange(numericValue, Integer.MIN_VALUE, Integer.MAX_VALUE, "Integer").intValue();
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                integerValue = parseIntegerValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                final boolean booleanCellValue = cell.getBooleanCellValue();
                integerValue = BooleanUtils.toInteger(booleanCellValue);
                break;

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return integerValue;
    }

    /**
     * Parses a string into an {@link Integer}. Trims first when {@code importTrim} is enabled and returns {@code null}
     * for blank input. When an import {@link DecimalFormat} is configured the whole string must match the pattern
     * ({@code PxlNumberSupport.parseFullyAsNumber}) and the parsed number is range-checked against the {@link Integer} range and
     * truncated; otherwise {@link Integer#parseInt(String)} is used.
     *
     * @param s          the source string
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@link Integer}, or {@code null} for blank input
     * @throws PxlCellCodecException if the string is not a valid {@link Integer} or is outside the {@link Integer} range
     */
    static Integer parseIntegerValue(final String s,
                                     final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        Integer integerValue;

        final DecimalFormat importDecimalFormatter = columnMeta.getImportDecimalFormatterCache();
        if (Objects.nonNull(importDecimalFormatter)) {
            integerValue = PxlNumberSupport.requireWithinRange(PxlNumberSupport.parseFullyAsNumber(importDecimalFormatter, stringValue, "Integer"), Integer.MIN_VALUE, Integer.MAX_VALUE, "Integer").intValue();
        } else {
            try {
                integerValue = Integer.parseInt(stringValue);
            } catch (NumberFormatException numberFormatException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "Integer"), numberFormatException);
            }
        }

        return integerValue;
    }

}
