package io.github.hclimkr.pxl.internal.codec;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.exception.PxlReflectionException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlExportColumnMeta;
import io.github.hclimkr.pxl.internal.meta.PxlImportColumnMeta;
import io.github.hclimkr.pxl.internal.support.PxlClassSupport;
import io.github.hclimkr.pxl.internal.support.PxlReflectionSupport;
import io.github.hclimkr.pxl.util.PxlCellUtils;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.*;
import java.util.*;

/**
 * Codec for {@link Collection} column values — splits a single cell into elements on import and joins
 * elements into one cell on export, delegating each element to the codec for the field's parameterized
 * type.
 *
 * <p>The column's import/export collection separator is treated as a whole-string literal (multi-character
 * separators are supported) and empty tokens are preserved to keep positional (index) fidelity;
 * {@code null} elements are written as empty strings on export.
 */
final class PxlCollectionCodec {

    /**
     * Prevents instantiation.
     */
    private PxlCollectionCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses the given cell into a collection by first reading it as a {@link String} via
     * {@link PxlStringCodec} and delegating to
     * {@link #parseCollectionValue(String, PxlImportColumnMeta)}.
     *
     * @param cell       the cell to read
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed collection, or {@code null} when the cell is empty
     * @throws PxlCellCodecException  if the cell type or element type is unsupported, or an element fails to parse
     * @throws PxlReflectionException if the element type cannot be resolved or the collection cannot be instantiated
     */
    static Collection<Object> parseCollectionValue(final Cell cell,
                                                   final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException, PxlReflectionException {

        final String stringValue = PxlStringCodec.parseStringValue(cell, columnMeta);

        return parseCollectionValue(stringValue, columnMeta);
    }

    /**
     * Parses a string into a collection by splitting on the column's import collection separator (a
     * whole-string literal, with empty tokens preserved for index fidelity) and parsing each token with the
     * codec for the field's parameterized element type. An empty input yields {@code null}; the concrete
     * collection type is derived from the column's declared collection class.
     *
     * @param collectionStringValue the joined string to split
     * @param columnMeta            the resolved import metadata for this column
     * @return the parsed collection, or {@code null} when empty
     * @throws PxlReflectionException if the element type cannot be resolved or the collection cannot be instantiated
     * @throws PxlCellCodecException  if the element type is unsupported or an element fails to parse
     */
    static Collection<Object> parseCollectionValue(final String collectionStringValue,
                                                   final PxlImportColumnMeta columnMeta)
            throws PxlReflectionException, PxlCellCodecException {

        final Class<?> collectionClass = columnMeta.getColumnClass();
        final Class<?> elementClass = PxlReflectionSupport.getParameterizedArgument0(columnMeta.getColumnField());

        if (StringUtils.isEmpty(collectionStringValue)) {
            return null;
        }

        final String importCollectionSeparator = columnMeta.getImportCollectionSeparator();
        // Treat the separator as a literal whole string (multi-character separators supported), and keep empty tokens to preserve positional (index) fidelity.
        // (split/splitPreserveAllTokens treat the separator as a 'set of characters', so a multi-character separator like "::" gets split on each character.)
        final String[] strings = StringUtils.splitByWholeSeparatorPreserveAllTokens(collectionStringValue, importCollectionSeparator);

        final Class<?> concreteCollectionClass = PxlClassSupport.getConcreteCollectionClass(collectionClass);
        final Collection<Object> valueObjects = (Collection<Object>) PxlReflectionSupport.newClassInstance(concreteCollectionClass);

        if (elementClass == String.class) {
            for (final String s : strings) {
                valueObjects.add(PxlStringCodec.parseStringValue(s, columnMeta));
            }
        } else if (elementClass == Byte.class) {
            for (final String s : strings) {
                valueObjects.add(PxlByteCodec.parseByteValue(s, columnMeta));
            }
        } else if (elementClass == Short.class) {
            for (final String s : strings) {
                valueObjects.add(PxlShortCodec.parseShortValue(s, columnMeta));
            }
        } else if (elementClass == Integer.class) {
            for (final String s : strings) {
                valueObjects.add(PxlIntegerCodec.parseIntegerValue(s, columnMeta));
            }
        } else if (elementClass == Long.class) {
            for (final String s : strings) {
                valueObjects.add(PxlLongCodec.parseLongValue(s, columnMeta));
            }
        } else if (elementClass == Double.class) {
            for (final String s : strings) {
                valueObjects.add(PxlDoubleCodec.parseDoubleValue(s, columnMeta));
            }
        } else if (elementClass == Float.class) {
            for (final String s : strings) {
                valueObjects.add(PxlFloatCodec.parseFloatValue(s, columnMeta));
            }
        } else if (elementClass == Character.class) {
            for (final String s : strings) {
                valueObjects.add(PxlCharacterCodec.parseCharacterValue(s, columnMeta));
            }
        } else if (elementClass == Boolean.class) {
            for (final String s : strings) {
                valueObjects.add(PxlBooleanCodec.parseBooleanValue(s, columnMeta));
            }
        } else if (elementClass == BigInteger.class) {
            for (final String s : strings) {
                valueObjects.add(PxlBigIntegerCodec.parseBigIntegerValue(s, columnMeta));
            }
        } else if (elementClass == BigDecimal.class) {
            for (final String s : strings) {
                valueObjects.add(PxlBigDecimalCodec.parseBigDecimalValue(s, columnMeta));
            }
        } else if (elementClass == Date.class) {
            for (final String s : strings) {
                valueObjects.add(PxlJavaDateCodec.parseJavaDateValue(s, columnMeta));
            }
        } else if (elementClass == LocalDate.class) {
            for (final String s : strings) {
                valueObjects.add(PxlLocalDateCodec.parseLocalDateValue(s, columnMeta));
            }
        } else if (elementClass == LocalTime.class) {
            for (final String s : strings) {
                valueObjects.add(PxlLocalTimeCodec.parseLocalTimeValue(s, columnMeta));
            }
        } else if (elementClass == LocalDateTime.class) {
            for (final String s : strings) {
                valueObjects.add(PxlLocalDateTimeCodec.parseLocalDateTimeValue(s, columnMeta));
            }
        } else if (elementClass == ZonedDateTime.class) {
            for (final String s : strings) {
                valueObjects.add(PxlZonedDateTimeCodec.parseZonedDateTimeValue(s, columnMeta));
            }
        } else if (elementClass == OffsetTime.class) {
            for (final String s : strings) {
                valueObjects.add(PxlOffsetTimeCodec.parseOffsetTimeValue(s, columnMeta));
            }
        } else if (elementClass == OffsetDateTime.class) {
            for (final String s : strings) {
                valueObjects.add(PxlOffsetDateTimeCodec.parseOffsetDateTimeValue(s, columnMeta));
            }
        } else if (elementClass == Duration.class) {
            for (final String s : strings) {
                valueObjects.add(PxlDurationCodec.parseDurationValue(s, columnMeta));
            }
        } else if (elementClass == Period.class) {
            for (final String s : strings) {
                valueObjects.add(PxlPeriodCodec.parsePeriodValue(s, columnMeta));
            }
        } else if (elementClass.isEnum()) {
            for (final String s : strings) {
                valueObjects.add(PxlEnumCodec.parseEnumValue(s, columnMeta));
            }
        } else if (columnMeta.isImportCustomConvertable()) {
            for (final String s : strings) {
                valueObjects.add(PxlObjectCodec.parseObjectValue(s, columnMeta));
            }
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_ELEMENT_TYPE_UNSUPPORTED, String.valueOf(elementClass.getSimpleName())));
        }

        return valueObjects;
    }

    /**
     * Writes a collection value into a single cell by exporting each element with the codec for the field's
     * parameterized element type and joining the results with the column's export collection separator.
     * {@code null} elements become empty strings (preserving index fidelity). A {@link String} source is
     * first split on the separator. When {@code exportStringAsPicture} is set for a {@link Collection}
     * source, elements are rendered as embedded pictures instead of text.
     *
     * @param cell       the target cell, or {@code null} to only compute the string
     * @param object     the source value (a {@link String} or {@link Collection})
     * @param columnMeta the resolved export metadata for this column
     * @return the joined string
     * @throws PxlReflectionException if the element type cannot be resolved or the collection cannot be instantiated
     * @throws PxlCellCodecException  if the source type or element type is unsupported, or an element fails to export
     * @throws PxlArgumentException   if a delegated element codec is given an invalid converter, styler, or unsupported target
     */
    static String buildCollectionCell(final Cell cell,
                                      final Object object,
                                      final PxlExportColumnMeta columnMeta)
            throws PxlReflectionException, PxlCellCodecException, PxlArgumentException {

        final Class<?> elementClass = PxlReflectionSupport.getParameterizedArgument0(columnMeta.getColumnField());

        Collection<?> collectionObject;
        if (object instanceof String) {
            final String collectionStringValue = (String) object;
            final String exportCollectionSeparator = columnMeta.getExportCollectionSeparator();
            final Object[] elementStrings = StringUtils.splitByWholeSeparatorPreserveAllTokens(collectionStringValue, exportCollectionSeparator);
            collectionObject = Arrays.asList(elementStrings);
        } else if (object instanceof Collection) {
            collectionObject = (Collection<?>) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "Collection"));
        }

        final List<String> strings = new ArrayList<>(PxlCollectionUtils.size(collectionObject));

        for (final Object element : collectionObject) {
            if (Objects.isNull(element)) {
                // To preserve positional (index) fidelity, null elements are written as empty strings.
                strings.add(StringUtils.EMPTY);
                continue;
            }

            if (elementClass == String.class) {
                strings.add(PxlStringCodec.buildStringCell(null, element, columnMeta));
            } else if (elementClass == Byte.class) {
                strings.add(PxlByteCodec.buildByteCell(null, element, columnMeta));
            } else if (elementClass == Short.class) {
                strings.add(PxlShortCodec.buildShortCell(null, element, columnMeta));
            } else if (elementClass == Integer.class) {
                strings.add(PxlIntegerCodec.buildIntegerCell(null, element, columnMeta));
            } else if (elementClass == Long.class) {
                strings.add(PxlLongCodec.buildLongCell(null, element, columnMeta));
            } else if (elementClass == Double.class) {
                strings.add(PxlDoubleCodec.buildDoubleCell(null, element, columnMeta));
            } else if (elementClass == Float.class) {
                strings.add(PxlFloatCodec.buildFloatCell(null, element, columnMeta));
            } else if (elementClass == Character.class) {
                strings.add(PxlCharacterCodec.buildCharacterCell(null, element, columnMeta));
            } else if (elementClass == Boolean.class) {
                strings.add(PxlBooleanCodec.buildBooleanCell(null, element, columnMeta));
            } else if (elementClass == BigInteger.class) {
                strings.add(PxlBigIntegerCodec.buildBigIntegerCell(null, element, columnMeta));
            } else if (elementClass == BigDecimal.class) {
                strings.add(PxlBigDecimalCodec.buildBigDecimalCell(null, element, columnMeta));
            } else if (elementClass == Date.class) {
                strings.add(PxlJavaDateCodec.buildJavaDateCell(null, element, columnMeta));
            } else if (elementClass == LocalDate.class) {
                strings.add(PxlLocalDateCodec.buildLocalDateCell(null, element, columnMeta));
            } else if (elementClass == LocalTime.class) {
                strings.add(PxlLocalTimeCodec.buildLocalTimeCell(null, element, columnMeta));
            } else if (elementClass == LocalDateTime.class) {
                strings.add(PxlLocalDateTimeCodec.buildLocalDateTimeCell(null, element, columnMeta));
            } else if (elementClass == ZonedDateTime.class) {
                strings.add(PxlZonedDateTimeCodec.buildZonedDateTimeCell(null, element, columnMeta));
            } else if (elementClass == OffsetTime.class) {
                strings.add(PxlOffsetTimeCodec.buildOffsetTimeCell(null, element, columnMeta));
            } else if (elementClass == OffsetDateTime.class) {
                strings.add(PxlOffsetDateTimeCodec.buildOffsetDateTimeCell(null, element, columnMeta));
            } else if (elementClass == Duration.class) {
                strings.add(PxlDurationCodec.buildDurationCell(null, element, columnMeta));
            } else if (elementClass == Period.class) {
                strings.add(PxlPeriodCodec.buildPeriodCell(null, element, columnMeta));
            } else if (elementClass.isEnum()) {
                strings.add(PxlEnumCodec.buildEnumCell(null, element, columnMeta));
            } else if (columnMeta.isExportCustomConvertable()) {
                strings.add(PxlObjectCodec.buildObjectCell(null, element, columnMeta));
            } else {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_ELEMENT_TYPE_UNSUPPORTED, String.valueOf(element.getClass().getSimpleName())));
            }
        }

        final String exportCollectionSeparator = columnMeta.getExportCollectionSeparator();
        final String cellString = StringUtils.join(strings, exportCollectionSeparator);

        if (Objects.nonNull(cell)) {
            if (object instanceof Collection && columnMeta.isExportStringAsPicture()) {
                PxlCellUtils.addPicturesToCell(cell,
                        strings,
                        PxlConstants.EXPORT_PICTURE_SCREEN_WIDTH_IN_PIXELS,
                        PxlConstants.EXPORT_PICTURE_SCREEN_HEIGHT_IN_PIXELS,
                        PxlConstants.EXPORT_PICTURE_SCREEN_PADDING_IN_PIXELS,
                        PxlConstants.EXPORT_HORIZONTAL_NUMBER_OF_PICTURE);
            } else {
                cell.setCellValue(cellString);
            }
        }

        return cellString;
    }

}
