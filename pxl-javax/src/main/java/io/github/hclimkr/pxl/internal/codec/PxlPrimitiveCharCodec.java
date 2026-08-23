package io.github.hclimkr.pxl.internal.codec;

import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlExportColumnMeta;
import io.github.hclimkr.pxl.internal.meta.PxlImportColumnMeta;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.NumberToTextConverter;

import java.util.Objects;
import java.util.Optional;

/**
 * Codec for primitive {@code char} column values - writes {@code char} into cells on export and parses cells and strings
 * into {@code char} on import. The character is taken as the first character of the cell's text (numeric cells are first rendered via
 * {@link NumberToTextConverter}); boolean cells map to {@code '1'}/{@code '0'}. Because {@code char} cannot be {@code null},
 * empty/blank input parses to {@code (char) 0}, the type's own default - the same choice the other primitive codecs make
 * ({@code 0}, {@code 0.0}), so an empty value leaves the field indistinguishable from one that was never set. Export always
 * writes the character as a text cell (no numeric formatting), after the column's export trim and masking options are
 * applied to it.
 */
final class PxlPrimitiveCharCodec {

    /**
     * Prevents instantiation.
     */
    private PxlPrimitiveCharCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Writes a {@code char} value into a cell. Accepts a {@link Character} directly or a {@link String} (its first character is
     * used; empty becomes {@code null}). A {@code null} value blanks the cell; otherwise the character goes through the
     * column's export trim and masking options and the resulting string is written as a text cell - trimming a whitespace
     * character therefore leaves an empty string.
     *
     * @param cell       the target cell (may be {@code null}, in which case only the return string is produced)
     * @param object     the source value ({@link Character} or {@link String})
     * @param columnMeta resolved export metadata for the column
     * @return the single-character string written to the cell, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the first character cannot be read or the object type cannot be converted to {@code char}
     */
    static String buildPrimitiveCharCell(final Cell cell,
                                         final Object object,
                                         final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Character charValue;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isEmpty(stringValue)) {
                charValue = null;
            } else {
                try {
                    charValue = stringValue.charAt(0);
                } catch (IndexOutOfBoundsException indexOutOfBoundsException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(stringValue), "char"), indexOutOfBoundsException);
                }
            }
        } else if (object instanceof Character) {
            charValue = (Character) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "char"));
        }

        if (Objects.isNull(charValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else {
            final String cellString = makePrimitiveCharExportString(charValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(cellString));
            return cellString;
        }
    }

    /**
     * Renders the export string for a {@code char}: takes the character's {@link Character#toString()} form and applies
     * string-level export processing via {@link PxlStringCodec#makeExportString}, or returns {@code null} when the value
     * is {@code null}.
     *
     * @param charValue  the value to render
     * @param columnMeta resolved export metadata for the column
     * @return the trimmed and/or masked single-character string, or {@code null} when the value is {@code null}
     */
    private static String makePrimitiveCharExportString(final Character charValue,
                                                        final PxlExportColumnMeta columnMeta) {

        if (Objects.isNull(charValue)) {
            return null;
        }

        return PxlStringCodec.makeExportString(charValue.toString(), columnMeta);
    }

    /**
     * Parses an Excel cell into a {@code char}. NUMERIC cells are rendered to text via {@link NumberToTextConverter} and the
     * first character is taken; STRING cells take their first character; BOOLEAN cells map to {@code '1'} (true) or {@code '0'}
     * (false); BLANK cells yield {@code (char) 0} (the resolver returns {@code null} for a blank cell before this codec is
     * reached, so the field keeps whatever it held).
     *
     * @param cell       the source cell
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@code char} ({@code (char) 0} for a blank cell)
     * @throws PxlCellCodecException if the cell type is unsupported
     */
    static char parsePrimitiveCharValue(final Cell cell,
                                        final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        char charValue = (char) 0;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                charValue = parsePrimitiveCharValue(NumberToTextConverter.toText(numericValue), columnMeta);
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                charValue = parsePrimitiveCharValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                final boolean booleanCellValue = cell.getBooleanCellValue();
                charValue = booleanCellValue ? '1' : '0';
                break;

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return charValue;
    }

    /**
     * Parses a string into a {@code char}. Trims first when {@code importTrim} is enabled and returns {@code (char) 0} for
     * empty input; otherwise returns the first character of the string (any remaining characters are ignored).
     *
     * @param s          the source string
     * @param columnMeta resolved import metadata for the column
     * @return the first character, or {@code (char) 0} for empty input
     * @throws PxlCellCodecException if the first character cannot be read
     */
    static char parsePrimitiveCharValue(final String s,
                                        final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isEmpty(stringValue)) {
            return (char) 0;
        }

        char charValue;

        try {
            charValue = stringValue.charAt(0);
        } catch (IndexOutOfBoundsException indexOutOfBoundsException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "char"), indexOutOfBoundsException);
        }

        return charValue;
    }

}
