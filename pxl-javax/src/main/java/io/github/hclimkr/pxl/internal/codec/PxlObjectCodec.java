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

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.Objects;
import java.util.Optional;

/**
 * Codec for custom object column values — parses strings into objects on import and writes objects into
 * cells on export, using the column's {@code @PxlImportConverter}/{@code @PxlExportConverter} method, a
 * {@code String} constructor, or a {@code toString} method. Blank values map to {@code null}.
 */
final class PxlObjectCodec {

    /**
     * Prevents instantiation.
     */
    private PxlObjectCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses the given cell into a custom object by first reading it as a {@code String} via
     * {@link PxlStringCodec} and delegating to {@link #parseObjectValue(String, PxlImportColumnMeta)}.
     *
     * @param cell       the cell to read
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed object, or {@code null} when blank
     * @throws PxlCellCodecException if the cell type is unsupported or the value cannot be converted
     */
    static Object parseObjectValue(final Cell cell,
                                   final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = PxlStringCodec.parseStringValue(cell, columnMeta);

        return parseObjectValue(stringValue, columnMeta);
    }

    /**
     * Parses a string token into a custom object using the column's import converter metadata (converter
     * method or {@code String} constructor). The value is trimmed when {@code importTrim} is set; a blank
     * value yields {@code null}.
     *
     * @param s          the raw string token
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed object, or {@code null} when blank
     * @throws PxlCellCodecException if no converter is available or conversion fails
     */
    static Object parseObjectValue(final String s,
                                   final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        return importStringToObject(stringValue, columnMeta.getImportCustomConverterMeta());
    }

    /**
     * Writes the given value as an object cell and returns the exported string. A {@code String} source is
     * imported and re-exported (sample flow); an instance of the converter's value class is used directly.
     * A {@code null} result blanks the cell.
     *
     * @param cell       the target cell, or {@code null} to only compute the string
     * @param object     the source value (a {@code String} or an instance of the converter's value class)
     * @param columnMeta the resolved export metadata for this column
     * @return the exported string, or {@code null} when blank
     * @throws PxlCellCodecException if the source type is unsupported or conversion fails
     * @throws PxlArgumentException  if the converter configuration is invalid
     */
    static String buildObjectCell(final Cell cell,
                                  final Object object,
                                  final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException, PxlArgumentException {

        final PxlExportConverterMeta converterMeta = columnMeta.getExportCustomConverterMeta();
        final Class<?> objectClass = converterMeta.getValueClass();

        Object objectValue;

        if (object instanceof String) {
            // The sample value is a string, so import it and then export it again.
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                objectValue = null;
            } else {
                objectValue = importStringToObject(stringValue, objectClass);
            }
        } else if (objectClass.isInstance(object)) {
            objectValue = object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "Object"));
        }

        if (Objects.isNull(objectValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else {
            final String cellString = makeObjectExportString(objectValue, columnMeta);
            Optional.ofNullable(cell).ifPresent(c -> c.setCellValue(cellString));
            return cellString;
        }
    }

    /**
     * Renders the export string for an object value: converts it via {@link #exportObjectToString} using the column's export
     * converter metadata, then applies string-level export processing via {@link PxlStringCodec#makeExportString}.
     *
     * @param objectValue the object value to render
     * @param columnMeta  resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value or the converter metadata is {@code null}
     * @throws PxlCellCodecException if the converter throws while producing the string
     */
    private static String makeObjectExportString(final Object objectValue,
                                                 final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final PxlExportConverterMeta converterMeta = columnMeta.getExportCustomConverterMeta();

        if (Objects.isNull(objectValue) || Objects.isNull(converterMeta)) {
            return null;
        }

        final String stringValue = exportObjectToString(objectValue, converterMeta);

        if (Objects.nonNull(stringValue)) {
            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } else {
            return stringValue;
        }
    }

    /**
     * Parses a string into an object of the given class, building import converter metadata from the class.
     *
     * @param stringValue the source string
     * @param objectClass the target object class
     * @return the parsed object, or {@code null} for blank input
     * @throws PxlCellCodecException if the string cannot be converted to the class
     * @throws PxlArgumentException  if converter metadata cannot be built from the class
     */
    private static Object importStringToObject(final String stringValue,
                                               final Class<?> objectClass)
            throws PxlCellCodecException, PxlArgumentException {

        return importStringToObject(stringValue, PxlImportColumnMeta.PxlImportConverterMeta.of(objectClass));
    }

    /**
     * Parses a string into an object using the given import converter metadata: a {@code @PxlImportConverter} method takes
     * precedence, otherwise a {@code String} constructor.
     *
     * @param stringValue   the source string
     * @param converterMeta the resolved import converter metadata
     * @return the parsed object, or {@code null} when the input is blank or the metadata is {@code null}
     * @throws PxlCellCodecException if neither a converter method nor a {@code String} constructor is available, or the conversion throws
     */
    private static Object importStringToObject(final String stringValue,
                                               final PxlImportColumnMeta.PxlImportConverterMeta converterMeta)
            throws PxlCellCodecException {

        if (StringUtils.isBlank(stringValue) || Objects.isNull(converterMeta)) {
            return null;
        }

        final Class<?> objectClass = converterMeta.getValueClass();
        final Method importConverterMethod = converterMeta.getImportConverterMethod();
        final Constructor<?> constructor = converterMeta.getStringConstructor();

        Object object = null;

        try {
            if (Objects.nonNull(importConverterMethod)) {
                object = importConverterMethod.invoke(null, stringValue);
            } else if (Objects.nonNull(constructor)) {
                object = constructor.newInstance(stringValue);
            } else {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_OBJECT_PARSE_FAILED, String.valueOf(stringValue), objectClass.getSimpleName()));
            }
        } catch (PxlCellCodecException e) {
            throw e;
        } catch (Exception e) {
            final Throwable cause = (e instanceof InvocationTargetException) ? e.getCause() : e;
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_OBJECT_PARSE_ERROR, String.valueOf(stringValue), objectClass.getSimpleName()), cause);
        }

        return object;
    }

    /**
     * Converts an object value to its export string using the given converter metadata: a {@code @PxlExportConverter} method
     * (static or instance) takes precedence, otherwise a {@code toString} method; returns {@code null} when neither is configured.
     *
     * @param objectValue   the object value to convert
     * @param converterMeta the resolved export converter metadata
     * @return the string form, or {@code null} when either argument is {@code null} or no converter is configured
     * @throws PxlCellCodecException if the converter throws while producing the string
     */
    private static String exportObjectToString(final Object objectValue,
                                               final PxlExportConverterMeta converterMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(objectValue) || Objects.isNull(converterMeta)) {
            return null;
        }

        final Class<?> objectClass = converterMeta.getValueClass();
        final Method exportConverterMethod = converterMeta.getExportConverterMethod();
        final Method toStringMethod = converterMeta.getToStringMethod();

        String stringValue = null;

        try {
            if (Objects.nonNull(exportConverterMethod)) {
                if (Modifier.isStatic(exportConverterMethod.getModifiers())) {
                    stringValue = (String) exportConverterMethod.invoke(null, objectValue);
                } else {
                    stringValue = (String) exportConverterMethod.invoke(objectValue);
                }
            } else if (Objects.nonNull(toStringMethod)) {
                stringValue = (String) toStringMethod.invoke(objectValue);
            }
        } catch (Exception e) {
            final Throwable cause = (e instanceof InvocationTargetException) ? e.getCause() : e;
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_OBJECT_FORMAT_ERROR, objectClass.getSimpleName()), cause);
        }

        return stringValue;
    }

}
