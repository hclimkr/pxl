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
 * Codec for primitive {@code float} column values - writes {@code float} into cells on export and parses cells and
 * strings into {@code float} on import. A numeric cell (a {@code double}) is narrowed to {@code float}; boolean cells map to
 * 1.0/0.0. Because {@code float} cannot be {@code null}, blank input parses to {@code 0.0f}. Both export and import reject
 * NaN and Infinity (import additionally rejects finite values that overflow the {@code float} range and narrow to Infinity),
 * keeping the two directions symmetric. To avoid float-to-double widening noise, export writes the value via
 * {@link PxlNumberSupport#floatToPlainDouble(float)} rendered as plain text by {@link NumberToTextConverter}.
 */
final class PxlPrimitiveFloatCodec {

    /**
     * Prevents instantiation.
     */
    private PxlPrimitiveFloatCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Writes a {@code float} value into a cell. Accepts a {@link Float} directly or a {@link String} (parsed via
     * {@link Float#parseFloat(String)}; blank becomes {@code null}). A {@code null} value blanks the cell; NaN or Infinity
     * is rejected. When the column is exported as text the value is formatted via {@link #makePrimitiveFloatExportString}
     * and written quote-prefixed; otherwise it is written as a numeric cell using {@link PxlNumberSupport#floatToPlainDouble(float)}.
     *
     * @param cell       the target cell (may be {@code null}, in which case only the return string is produced)
     * @param object     the source value ({@link Float} or {@link String})
     * @param columnMeta resolved export metadata for the column
     * @return the string representation of the written value, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if a string value is malformed, the object type cannot be converted to {@code float},
     *                               or the value is NaN or Infinity
     */
    static String buildPrimitiveFloatCell(final Cell cell,
                                          final Object object,
                                          final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Float floatValue;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                floatValue = null;
            } else {
                try {
                    floatValue = Float.parseFloat(stringValue);
                } catch (NumberFormatException numberFormatException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(stringValue), "float"), numberFormatException);
                }
            }
        } else if (object instanceof Float) {
            floatValue = (Float) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "float"));
        }

        if (Objects.isNull(floatValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        }

        PxlNumberSupport.requireFiniteForExport(floatValue);

        if (columnMeta.isExportedToString()) {
            final String cellString = makePrimitiveFloatExportString(floatValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> columnMeta.setQuotePrefixedCellValue(c, cellString));
            return cellString;
        } else {
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(PxlNumberSupport.floatToPlainDouble(floatValue)));
            return NumberToTextConverter.toText(PxlNumberSupport.floatToPlainDouble(floatValue));
        }
    }

    /**
     * Renders the export string for a {@code float}: applies the export {@link DecimalFormat} when configured, otherwise
     * applies masking when an export masking pattern is set, otherwise returns plain text; in all branches the value first
     * passes through {@link PxlNumberSupport#floatToPlainDouble(float)} to avoid widening noise.
     *
     * @param floatValue the value to render
     * @param columnMeta resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the export pattern cannot be applied to the value
     */
    private static String makePrimitiveFloatExportString(final Float floatValue,
                                                         final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(floatValue)) {
            return null;
        }

        final DecimalFormat exportDecimalFormatter = columnMeta.getExportDecimalFormatterCache();
        final Pattern exportMaskingPattern = columnMeta.getExportMaskingPattern();

        if (Objects.nonNull(exportDecimalFormatter)) {
            try {
                final String stringValue = exportDecimalFormatter.format(PxlNumberSupport.floatToPlainDouble(floatValue));
                return PxlStringCodec.makeExportString(stringValue, columnMeta);
            } catch (IllegalArgumentException illegalArgumentException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PATTERN_APPLY_FAILED, String.valueOf(floatValue)), illegalArgumentException);
            }
        } else if (Objects.nonNull(exportMaskingPattern)) {
            final String stringValue = NumberToTextConverter.toText(PxlNumberSupport.floatToPlainDouble(floatValue));
            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } else {
            return NumberToTextConverter.toText(PxlNumberSupport.floatToPlainDouble(floatValue));
        }
    }

    /**
     * Parses an Excel cell into a {@code float}. NUMERIC cells are narrowed from {@code double} to {@code float}; STRING
     * cells are delegated to the string overload; BOOLEAN cells map to 1.0f (true) or 0.0f (false); BLANK cells yield {@code 0.0f}.
     *
     * @param cell       the source cell
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@code float} (0.0f for a blank cell)
     * @throws PxlCellCodecException if the cell type is unsupported
     */
    static float parsePrimitiveFloatValue(final Cell cell,
                                          final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        float floatValue = 0.f;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                floatValue = (float) numericValue;
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                floatValue = parsePrimitiveFloatValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                final boolean booleanCellValue = cell.getBooleanCellValue();
                floatValue = (float) BooleanUtils.toInteger(booleanCellValue);
                break;

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        PxlNumberSupport.requireFiniteForImport(floatValue, "float");

        return floatValue;
    }

    /**
     * Parses a string into a {@code float}. Trims first when {@code importTrim} is enabled and returns {@code 0.0f} for blank
     * input. When an import {@link DecimalFormat} is configured the whole string must match the pattern
     * ({@code PxlNumberSupport.parseFullyAsNumber}) and its parsed value is narrowed to {@code float}; otherwise
     * {@link Float#parseFloat(String)} is used.
     *
     * @param s          the source string
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@code float} (0.0f for blank input)
     * @throws PxlCellCodecException if the string is not a valid {@code float}
     */
    static float parsePrimitiveFloatValue(final String s,
                                          final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return (float) 0.f;
        }

        float floatValue;

        final DecimalFormat importDecimalFormatter = columnMeta.getImportDecimalFormatterCache();
        if (Objects.nonNull(importDecimalFormatter)) {
            floatValue = PxlNumberSupport.parseFullyAsNumber(importDecimalFormatter, stringValue, "float").floatValue();
        } else {
            try {
                floatValue = Float.parseFloat(stringValue);
            } catch (NumberFormatException numberFormatException) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "float"), numberFormatException);
            }
        }

        PxlNumberSupport.requireFiniteForImport(floatValue, "float");

        return floatValue;
    }

}
