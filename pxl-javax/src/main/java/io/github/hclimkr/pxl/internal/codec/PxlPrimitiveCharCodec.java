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
 * Codec for primitive {@code char} column values — parses cells and strings into {@code char} on import and writes {@code char}
 * into cells on export. The character is taken as the first character of the cell's text (numeric cells are first rendered via
 * {@link NumberToTextConverter}); boolean cells map to {@code '1'}/{@code '0'}. Because {@code char} cannot be {@code null},
 * empty/blank input parses to a space {@code ' '}. Export always writes the character as a text cell (no numeric formatting or masking).
 */
final class PxlPrimitiveCharCodec {

    /**
     * Prevents instantiation.
     */
    private PxlPrimitiveCharCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses an Excel cell into a {@code char}. NUMERIC cells are rendered to text via {@link NumberToTextConverter} and the
     * first character is taken; STRING cells take their first character; BOOLEAN cells map to {@code '1'} (true) or {@code '0'}
     * (false); BLANK cells yield a space {@code ' '}.
     *
     * @param cell       the source cell
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@code char} (a space for a blank cell)
     * @throws PxlCellCodecException if the cell type is unsupported
     */
    static char parsePrimitiveCharValue(final Cell cell,
                                        final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        char charValue = ' ';

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
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return charValue;
    }

    /**
     * Parses a string into a {@code char}. Trims first when {@code importTrim} is enabled and returns a space {@code ' '} for
     * empty input; otherwise returns the first character of the string (any remaining characters are ignored).
     *
     * @param s          the source string
     * @param columnMeta resolved import metadata for the column
     * @return the first character, or a space for empty input
     * @throws PxlCellCodecException if the first character cannot be read
     */
    static char parsePrimitiveCharValue(final String s,
                                        final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isEmpty(stringValue)) {
            return ' ';
        }

        char charValue;

        try {
            charValue = stringValue.charAt(0);
        } catch (IndexOutOfBoundsException indexOutOfBoundsException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "char"), indexOutOfBoundsException);
        }

        return charValue;
    }

    /**
     * Writes a {@code char} value into a cell. Accepts a {@code Character} directly or a {@code String} (its first character is
     * used; empty becomes {@code null}). A {@code null} value blanks the cell; otherwise the single-character string is written
     * as a text cell.
     *
     * @param cell       the target cell (may be {@code null}, in which case only the return string is produced)
     * @param object     the source value ({@code Character} or {@code String})
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
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "char"), indexOutOfBoundsException);
                }
            }
        } else if (object instanceof Character) {
            charValue = (Character) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "char"));
        }

        if (Objects.isNull(charValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else {
            final String cellString = makePrimitiveCharExportString(charValue);
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(cellString));
            return cellString;
        }
    }

    /**
     * Renders the export string for a {@code char}: returns the character's {@link Character#toString()} form, or {@code null}
     * when the value is {@code null}.
     *
     * @param charValue the value to render
     * @return the single-character string, or {@code null} when the value is {@code null}
     */
    private static String makePrimitiveCharExportString(final Character charValue) {

        if (Objects.isNull(charValue)) {
            return null;
        }

        return charValue.toString();
    }

}
