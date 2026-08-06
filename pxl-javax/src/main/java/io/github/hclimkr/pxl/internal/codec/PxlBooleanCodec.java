package io.github.hclimkr.pxl.internal.codec;

import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlExportColumnMeta;
import io.github.hclimkr.pxl.internal.meta.PxlImportColumnMeta;
import org.apache.commons.lang3.BooleanUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.util.Objects;
import java.util.Optional;

/**
 * Codec for {@link Boolean} column values — parses cells/strings into {@link Boolean} on import and
 * writes {@link Boolean} into cells on export.
 *
 * <p>NUMERIC cells are {@code true} when non-zero; STRING cells honour the column's import true/false
 * strings (matched case-insensitively) before falling back to {@link BooleanUtils#toBooleanObject};
 * BLANK/blank values map to {@code null}. Export renders the value with the column's export
 * true/false/null strings.
 */
final class PxlBooleanCodec {

    /**
     * Prevents instantiation.
     */
    private PxlBooleanCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses the given cell into a {@link Boolean}. NUMERIC cells are {@code true} when their absolute
     * value exceeds 1e-7; STRING cells are delegated to
     * {@link #parseBooleanValue(String, PxlImportColumnMeta)}; BOOLEAN cells are returned directly; BLANK
     * cells yield {@code null}.
     *
     * @param cell       the cell to read
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link Boolean}, or {@code null} when blank
     * @throws PxlCellCodecException if the cell type is not supported
     */
    static Boolean parseBooleanValue(final Cell cell,
                                     final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Boolean booleanValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                booleanValue = Math.abs(numericValue) > 0.0000001;
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                booleanValue = parseBooleanValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                booleanValue = cell.getBooleanCellValue();
                break;

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return booleanValue;
    }

    /**
     * Parses a string token into a {@link Boolean}. The value is trimmed when {@code importTrim} is set; a
     * blank value yields {@code null}. It matches the column's import true/false strings case-insensitively,
     * then falls back to {@link BooleanUtils#toBooleanObject}.
     *
     * @param s          the raw string token
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link Boolean}, or {@code null} when blank
     * @throws PxlCellCodecException if the value is not a recognizable boolean
     */
    static Boolean parseBooleanValue(final String s,
                                     final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        final String importTrueString = columnMeta.getImportTrueString();
        if (StringUtils.isNotBlank(importTrueString) && StringUtils.equalsIgnoreCase(stringValue, importTrueString)) {
            return Boolean.TRUE;
        }

        final String importFalseString = columnMeta.getImportFalseString();
        if (StringUtils.isNotBlank(importFalseString) && StringUtils.equalsIgnoreCase(stringValue, importFalseString)) {
            return Boolean.FALSE;
        }

        Boolean booleanValue = BooleanUtils.toBooleanObject(stringValue);
        if (Objects.isNull(booleanValue)) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "Boolean"));
        }

        return booleanValue;
    }

    /**
     * Writes the given value as a {@link Boolean} cell and returns the exported string. A {@link String}
     * source is parsed via {@link BooleanUtils#toBooleanObject}; a {@link Boolean} source is used directly.
     * A {@code null} result blanks the cell; otherwise the column's export true/false/null strings are
     * written.
     *
     * @param cell       the target cell, or {@code null} to only compute the string
     * @param object     the source value (a {@link String} or {@link Boolean})
     * @param columnMeta the resolved export metadata for this column
     * @return the exported string, or {@code null} when the value is blank
     * @throws PxlCellCodecException if the source is not a {@link String}/{@link Boolean}, or an invalid
     *                               boolean string
     */
    static String buildBooleanCell(final Cell cell,
                                   final Object object,
                                   final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Boolean booleanValue;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                booleanValue = null;
            } else {
                booleanValue = BooleanUtils.toBooleanObject(stringValue);
                if (Objects.isNull(booleanValue)) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(stringValue), "Boolean"));
                }
            }
        } else if (object instanceof Boolean) {
            booleanValue = (Boolean) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "Boolean"));
        }

        if (Objects.isNull(booleanValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else {
            final String cellString = makeBooleanExportString(booleanValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(cellString));
            return cellString;
        }
    }

    /**
     * Renders the export string for a {@link Boolean} using the configured true/false/null string representations.
     *
     * @param booleanValue the value to render
     * @param columnMeta   resolved export metadata for the column
     * @return the configured true or false string, or the configured null string when the value is {@code null}
     */
    private static String makeBooleanExportString(final Boolean booleanValue,
                                                  final PxlExportColumnMeta columnMeta) {

        final String exportNullString = columnMeta.getExportNullString();
        final String exportTrueString = columnMeta.getExportTrueString();
        final String exportFalseString = columnMeta.getExportFalseString();

        return BooleanUtils.toString(booleanValue, exportTrueString, exportFalseString, exportNullString);
    }

}
