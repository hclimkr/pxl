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
 * Codec for {@link Character} column values - writes {@link Character} into cells on export and parses
 * cells/strings into {@link Character} on import. The first character of the cell/string is taken; NUMERIC
 * cells are stringified via {@link NumberToTextConverter} and BOOLEAN cells map to {@code '1'}/{@code '0'}.
 */
final class PxlCharacterCodec {

    /**
     * Prevents instantiation.
     */
    private PxlCharacterCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Writes the given value as a single-character string cell and returns it. A {@link String} source
     * contributes its first character; a {@link Character} source is used directly. A {@code null} result
     * blanks the cell.
     *
     * @param cell       the target cell, or {@code null} to only compute the string
     * @param object     the source value (a {@link String} or {@link Character})
     * @param columnMeta the resolved export metadata for this column
     * @return the exported one-character string, or {@code null} when empty
     * @throws PxlCellCodecException if the source is not a {@link String}/{@link Character}
     */
    static String buildCharacterCell(final Cell cell,
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
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(stringValue), "Character"), indexOutOfBoundsException);
                }
            }
        } else if (object instanceof Character) {
            charValue = (Character) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "Character"));
        }

        if (Objects.isNull(charValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else {
            final String cellString = makeCharacterExportString(charValue);
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(cellString));
            return cellString;
        }
    }

    /**
     * Renders the export string for a {@link Character}: its {@link Character#toString()} form.
     *
     * @param charValue the value to render
     * @return the single-character string, or {@code null} when the value is {@code null}
     */
    private static String makeCharacterExportString(final Character charValue) {

        if (Objects.isNull(charValue)) {
            return null;
        }

        return charValue.toString();
    }

    /**
     * Parses the given cell into a {@link Character}. NUMERIC cells are converted to text with
     * {@link NumberToTextConverter} and their first character taken; STRING cells take their first
     * character; BOOLEAN cells map to {@code '1'} (true) or {@code '0'} (false); BLANK cells yield
     * {@code null}.
     *
     * @param cell       the cell to read
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link Character}, or {@code null} when blank
     * @throws PxlCellCodecException if the cell type is not supported
     */
    static Character parseCharacterValue(final Cell cell,
                                         final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Character charValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                charValue = parseCharacterValue(NumberToTextConverter.toText(numericValue), columnMeta);
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                charValue = parseCharacterValue(stringCellValue, columnMeta);
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
     * Parses a string token into a {@link Character} by taking its first character. The value is trimmed
     * when {@code importTrim} is set; an empty value yields {@code null}.
     *
     * @param s          the raw string token
     * @param columnMeta the resolved import metadata for this column
     * @return the first character, or {@code null} when empty
     * @throws PxlCellCodecException if the character cannot be read
     */
    static Character parseCharacterValue(final String s,
                                         final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isEmpty(stringValue)) {
            return null;
        }

        Character charValue;

        try {
            charValue = stringValue.charAt(0);
        } catch (IndexOutOfBoundsException indexOutOfBoundsException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "Character"), indexOutOfBoundsException);
        }

        return charValue;
    }

}
