package io.github.hclimkr.pxl.internal.codec;

import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlExportColumnMeta;
import io.github.hclimkr.pxl.internal.meta.PxlExportColumnMeta.PxlExportConverterMeta;
import io.github.hclimkr.pxl.internal.meta.PxlImportColumnMeta;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.Arrays;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Stream;

/**
 * Codec for Java {@code enum} column values — parses cells/strings into enum constants on import and
 * writes enum constants into cells on export.
 *
 * <p>Conversion honours any {@code @PxlImportConverter}/{@code @PxlExportConverter} method or a
 * {@link String} constructor; otherwise it matches a constant by its {@code toString} result (falling back
 * to {@link Enum#name()}), comparing case-insensitively and ignoring whitespace. BLANK/blank values map to
 * {@code null}.
 */
public final class PxlEnumCodec {

    /**
     * Prevents instantiation.
     */
    private PxlEnumCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses the given cell into an enum constant. NUMERIC cells are first stringified via
     * {@link PxlStringCodec}; STRING cells are read directly; BOOLEAN cells are rejected as an unsupported
     * cell type; BLANK cells yield {@code null}.
     *
     * @param cell       the cell to read
     * @param columnMeta the resolved import metadata for this column
     * @return the matching enum constant, or {@code null} when blank
     * @throws PxlCellCodecException if the cell type is unsupported or the value cannot be converted
     */
    static Object parseEnumValue(final Cell cell,
                                 final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        String cellValue = null;
        Object enumValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                cellValue = PxlStringCodec.parseStringValue(cell, columnMeta);
                enumValue = parseEnumValue(cellValue, columnMeta);
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                enumValue = parseEnumValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return enumValue;
    }

    /**
     * Parses a string token into an enum constant using the column's import converter metadata. The value
     * is trimmed when {@code importTrim} is set; a blank value yields {@code null}.
     *
     * @param s          the raw string token
     * @param columnMeta the resolved import metadata for this column
     * @return the matching enum constant, or {@code null} when blank
     * @throws PxlCellCodecException if the value cannot be converted to the enum type
     */
    static Object parseEnumValue(final String s,
                                 final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        return importStringToEnum(stringValue, columnMeta.getImportCustomConverterMeta());
    }

    /**
     * Writes the given value as an enum cell and returns the exported string. A {@link String} source is
     * imported to an enum and re-exported (sample flow); an {@link Enum} source is matched to the target
     * enum by ordinal. A {@code null} result blanks the cell.
     *
     * @param cell       the target cell, or {@code null} to only compute the string
     * @param object     the source value (a {@link String} or {@link Enum})
     * @param columnMeta the resolved export metadata for this column
     * @return the exported string, or {@code null} when blank
     * @throws PxlCellCodecException if the source is not a {@link String}/{@link Enum}, or the value cannot be converted
     * @throws PxlArgumentException  if the resolved converter or target enum type is invalid
     */
    static String buildEnumCell(final Cell cell,
                                final Object object,
                                final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException, PxlArgumentException {

        final PxlExportConverterMeta converterMeta = columnMeta.getExportCustomConverterMeta();
        final Class<?> enumClass = converterMeta.getValueClass();

        Object enumValue;

        if (object instanceof String) {
            // The sample value is a string, so import it and then export it again.
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                enumValue = null;
            } else {
                enumValue = importStringToEnum(stringValue, enumClass);
            }
        } else if (object instanceof Enum) {
            final Enum<?> enumVariable = (Enum<?>) object;

            enumValue = Arrays.stream(enumClass.getEnumConstants())
                    .filter(o -> ((Enum<?>) o).ordinal() == enumVariable.ordinal())
                    .findFirst()
                    .orElse(null);

            if (Objects.isNull(enumValue)) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "enum"));
            }
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "enum"));
        }

        if (Objects.isNull(enumValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else {
            final String cellString = makeEnumExportString(enumValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(cellString));
            return cellString;
        }
    }

    /**
     * Renders the export string for an enum value: converts it via {@link #exportEnumToString} using the column's export
     * converter metadata, then applies string-level export processing via {@link PxlStringCodec#makeExportString}.
     *
     * @param enumValue  the enum value to render
     * @param columnMeta resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value or the converter metadata is {@code null}
     * @throws PxlCellCodecException if the converter throws while producing the string
     */
    private static String makeEnumExportString(final Object enumValue,
                                               final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final PxlExportConverterMeta converterMeta = columnMeta.getExportCustomConverterMeta();

        if (Objects.isNull(enumValue) || Objects.isNull(converterMeta)) {
            return null;
        }

        final String stringValue = exportEnumToString(enumValue, converterMeta);

        if (Objects.nonNull(stringValue)) {
            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } else {
            return stringValue;
        }
    }

    /**
     * Parses a string into an enum constant of the given class, building import converter metadata from the class.
     *
     * @param stringValue the source string
     * @param enumClass   the target enum class
     * @return the parsed enum constant, or {@code null} for blank input
     * @throws PxlCellCodecException if the string cannot be resolved to a constant of the class
     * @throws PxlArgumentException  if converter metadata cannot be built from the class
     */
    private static Object importStringToEnum(final String stringValue,
                                             final Class<?> enumClass)
            throws PxlCellCodecException, PxlArgumentException {

        return importStringToEnum(stringValue, PxlImportColumnMeta.PxlImportConverterMeta.of(enumClass));
    }

    /**
     * Parses a string into an enum constant using the given import converter metadata: a {@code @PxlImportConverter} method
     * takes precedence, then a {@link String} constructor, then matching an enum constant by its {@code toString} value and
     * finally by {@link Enum#name()} (both compared case- and whitespace-insensitively).
     *
     * @param stringValue   the source string
     * @param converterMeta the resolved import converter metadata
     * @return the parsed enum constant, or {@code null} when the input is blank or the metadata is {@code null}
     * @throws PxlCellCodecException if no matching constant is found or the converter throws while parsing
     */
    private static Object importStringToEnum(final String stringValue,
                                             final PxlImportColumnMeta.PxlImportConverterMeta converterMeta)
            throws PxlCellCodecException {

        if (StringUtils.isBlank(stringValue) || Objects.isNull(converterMeta)) {
            return null;
        }

        final Class<?> enumClass = converterMeta.getValueClass();
        final Method importConverterMethod = converterMeta.getImportConverterMethod();
        final Constructor<?> constructor = converterMeta.getStringConstructor();
        final Method toStringMethod = converterMeta.getToStringMethod();

        Object object = null;

        try {
            if (Objects.nonNull(importConverterMethod)) {
                object = importConverterMethod.invoke(null, stringValue);
            } else if (Objects.nonNull(constructor)) {
                object = constructor.newInstance(stringValue);
            } else if (Objects.nonNull(toStringMethod)) {
                object = Stream.of(enumClass.getEnumConstants())
                        .filter(o -> {
                            try {
                                return equalsEnumString((String) toStringMethod.invoke(o), stringValue);
                            } catch (ReflectiveOperationException ignored) {
                                return false;
                            }
                        })
                        .findFirst()
                        .orElse(null);

                if (Objects.isNull(object)) {
                    object = Stream.of(enumClass.getEnumConstants())
                            .filter(o -> equalsEnumString(((Enum<?>) o).name(), stringValue))
                            .findFirst()
                            .orElse(null);

                    if (Objects.isNull(object)) {
                        throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_ENUM_PARSE_FAILED, String.valueOf(stringValue), enumClass.getSimpleName()));
                    }
                }
            }
        } catch (PxlCellCodecException e) {
            throw e;
        } catch (Exception e) {
            final Throwable cause = (e instanceof InvocationTargetException) ? e.getCause() : e;
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_ENUM_PARSE_ERROR, String.valueOf(stringValue), enumClass.getSimpleName()), cause);
        }

        return object;
    }

    /**
     * Converts an enum value to its export string using the given converter metadata: a
     * {@code @PxlExportConverter} method (static or instance) takes precedence, then a {@code toString}
     * method, otherwise the constant's {@link Enum#name()}.
     *
     * @param enumValue     the enum value to convert
     * @param converterMeta the resolved export converter metadata
     * @return the string form, or {@code null} when either argument is {@code null}
     * @throws PxlCellCodecException if the converter throws while producing the string
     */
    public static String exportEnumToString(final Object enumValue,
                                            final PxlExportConverterMeta converterMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(enumValue) || Objects.isNull(converterMeta)) {
            return null;
        }

        final Class<?> enumClass = converterMeta.getValueClass();
        final Method exportConverterMethod = converterMeta.getExportConverterMethod();
        final Method toStringMethod = converterMeta.getToStringMethod();

        String stringValue;

        try {
            if (Objects.nonNull(exportConverterMethod)) {
                if (Modifier.isStatic(exportConverterMethod.getModifiers())) {
                    stringValue = (String) exportConverterMethod.invoke(null, enumValue);
                } else {
                    stringValue = (String) exportConverterMethod.invoke(enumValue);
                }
                return stringValue;
            } else if (Objects.nonNull(toStringMethod)) {
                stringValue = (String) toStringMethod.invoke(enumValue);
            } else {
                stringValue = ((Enum<?>) enumValue).name();
            }
        } catch (Exception e) {
            final Throwable cause = (e instanceof InvocationTargetException) ? e.getCause() : e;
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_ENUM_FORMAT_ERROR, enumClass.getSimpleName()), cause);
        }

        return stringValue;
    }

    /**
     * Compares two enum-related strings for equality, ignoring case and all whitespace; the same reference (including both
     * {@code null}) is equal, while a single {@code null} is not.
     *
     * @param s1 the first string
     * @param s2 the second string
     * @return {@code true} if the two strings are considered equal
     */
    private static boolean equalsEnumString(final String s1,
                                            final String s2) {

        if (s1 == s2) {
            return true;
        }

        if (Objects.isNull(s1) || Objects.isNull(s2)) {
            return false;
        }

        return StringUtils.equalsIgnoreCase(StringUtils.deleteWhitespace(s1), StringUtils.deleteWhitespace(s2));
    }

}
