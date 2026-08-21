package io.github.hclimkr.pxl.internal.codec;

import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlExportColumnMeta;
import io.github.hclimkr.pxl.internal.meta.PxlImportColumnMeta;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.util.Objects;
import java.util.Optional;
import java.util.UUID;
import java.util.regex.Pattern;

/**
 * Codec for {@link UUID} column values - writes {@link UUID} into cells on export and parses cells and
 * strings into {@link UUID} on import. A UUID only has meaning as text, so it is always written as a string cell in its
 * canonical lower-case 8-4-4-4-12 form ({@link UUID#toString()}).
 *
 * <p>Import accepts the canonical form only, in either case. The check is made here rather than left to
 * {@link UUID#fromString(String)}, which counts the hyphen-separated groups but not their digits and so widens
 * {@code "1-1-1-1-1"} into a completely different UUID; how lenient it is also depends on the JDK the library runs on.
 * Validating the form first keeps a typo from silently becoming another value and keeps the same file reading the same
 * way on every runtime, in the spirit of the non-lenient date parsing elsewhere in the library. A hyphen-less 32-digit
 * form, a braced {@code "{...}"} form and a {@code "urn:uuid:"} prefix are all rejected, since export only ever writes
 * the canonical form.
 *
 * <p>The column's {@code exportPattern}/{@code importPattern} carry no meaning for a UUID and are ignored;
 * {@code exportTrim} and {@code exportMasking} apply to the canonical string as they do for any other type.
 */
final class PxlUuidCodec {

    /**
     * The canonical 8-4-4-4-12 hexadecimal form, matched in either case.
     */
    private static final Pattern CANONICAL_UUID_PATTERN =
            Pattern.compile("[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}");

    /**
     * Prevents instantiation.
     */
    private PxlUuidCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Writes a {@link UUID} value into a cell. Accepts a {@link UUID} directly or a {@link String} (the sample export
     * value, which is parsed by the same rules as an imported one; blank becomes {@code null}). A {@code null} value
     * blanks the cell. The value is always written as text in its canonical form, after the column's export trim and
     * masking are applied.
     *
     * @param cell       the target cell (may be {@code null}, in which case only the return string is produced)
     * @param object     the source value ({@link UUID} or {@link String})
     * @param columnMeta resolved export metadata for the column
     * @return the string representation of the written value, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if a string value is not a canonical UUID or the object type cannot be converted to {@link UUID}
     */
    static String buildUuidCell(final Cell cell,
                                final Object object,
                                final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        UUID uuidValue;

        if (object instanceof String) {
            // The sample value is a string, so parse it and then export it again.
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                uuidValue = null;
            } else {
                uuidValue = toUuid(stringValue, true);
            }
        } else if (object instanceof UUID) {
            uuidValue = (UUID) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "UUID"));
        }

        if (Objects.isNull(uuidValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else {
            final String cellString = makeUuidExportString(uuidValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(cellString));
            return cellString;
        }
    }

    /**
     * Renders the export string for a {@link UUID}: its canonical lower-case form, with the column's export trim and
     * masking applied over it. There is no numeric or date form to fall back on, so no export pattern is consulted.
     *
     * @param uuidValue  the value to render
     * @param columnMeta resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     */
    private static String makeUuidExportString(final UUID uuidValue,
                                               final PxlExportColumnMeta columnMeta) {

        if (Objects.isNull(uuidValue)) {
            return null;
        }

        return PxlStringCodec.makeExportString(uuidValue.toString(), columnMeta);
    }

    /**
     * Parses an Excel cell into a {@link UUID}. STRING cells are delegated to the string overload; NUMERIC cells are
     * first rendered to text via {@link PxlStringCodec} and then parsed the same way, so a cell that carries a number
     * by mistake is reported as an invalid value rather than an unsupported cell type; BOOLEAN cells are rejected;
     * BLANK cells yield {@code null}.
     *
     * @param cell       the source cell
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@link UUID}, or {@code null} for a blank cell
     * @throws PxlCellCodecException if the cell type is unsupported or the value is not a canonical UUID
     */
    static UUID parseUuidValue(final Cell cell,
                               final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        UUID uuidValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final String numericStringValue = PxlStringCodec.parseStringValue(cell, columnMeta);
                uuidValue = parseUuidValue(numericStringValue, columnMeta);
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                uuidValue = parseUuidValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return uuidValue;
    }

    /**
     * Parses a string into a {@link UUID}. Trims first when {@code importTrim} is enabled and returns {@code null} for
     * blank input; the remaining value must be a canonical UUID.
     *
     * @param s          the source string
     * @param columnMeta resolved import metadata for the column
     * @return the parsed {@link UUID}, or {@code null} for blank input
     * @throws PxlCellCodecException if the string is not a canonical UUID
     */
    static UUID parseUuidValue(final String s,
                               final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        return toUuid(stringValue, false);
    }

    /**
     * Converts a non-blank string to a {@link UUID}, reporting failures with the diagnostic keys of the calling
     * direction. Both directions parse strings: export resolves a string sample value before writing it back out,
     * and import reads a cell. The string must match the canonical form before {@link UUID#fromString(String)} is asked to
     * build the value.
     *
     * @param stringValue the source string
     * @param forExport   {@code true} when called from the export path, {@code false} from the import path
     * @return the parsed {@link UUID}
     * @throws PxlCellCodecException if the string is not a canonical UUID
     */
    private static UUID toUuid(final String stringValue,
                               final boolean forExport)
            throws PxlCellCodecException {

        if (!CANONICAL_UUID_PATTERN.matcher(stringValue).matches()) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(forExport ?
                    PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID :
                    PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "UUID"));
        }

        try {
            return UUID.fromString(stringValue);
        } catch (IllegalArgumentException illegalArgumentException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(forExport ?
                    PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID :
                    PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "UUID"), illegalArgumentException);
        }
    }

}
