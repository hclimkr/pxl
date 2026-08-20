package io.github.hclimkr.pxl.internal.codec;

import com.github.pjfanning.xlsx.impl.StreamingCell;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.exception.PxlReflectionException;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlExportColumnMeta;
import io.github.hclimkr.pxl.internal.meta.PxlImportColumnMeta;
import io.github.hclimkr.pxl.internal.support.PxlClassSupport;
import io.github.hclimkr.pxl.util.PxlCellUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.*;
import java.util.Date;
import java.util.Objects;
import java.util.Optional;
import java.util.UUID;

/**
 * Central type dispatcher for cell value conversion. Routes each column, based on its resolved
 * {@code columnClass}, to the matching per-type codec for both import (cell/string to value) and export
 * (value to cell). Covers {@link String}, the primitive and boxed numeric types, {@code char}/{@link Character},
 * {@code boolean}/{@link Boolean}, {@link BigInteger}/{@link BigDecimal}, the {@code java.time} types plus
 * {@link Date}, {@link UUID}, enums, collections, and objects handled by a custom converter.
 */
public final class PxlCellResolver {

    /**
     * Prevents instantiation.
     */
    private PxlCellResolver() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Reads a single Excel cell and converts it to the column's declared type. Returns {@code null} when the cell,
     * {@code columnMeta}, or the resolved field/class is {@code null}, or when the cell is blank. Formula cells are
     * evaluated in place via {@code formulaEvaluator} before dispatch; formulas cannot be read through the streaming
     * reader. Dispatches on {@code columnMeta.getColumnClass()} to the matching per-type codec.
     *
     * @param cell             the source cell (may be {@code null})
     * @param columnMeta       resolved import metadata for the target column (may be {@code null})
     * @param formulaEvaluator evaluator used to resolve formula cells; must be non-null when a formula cell is encountered
     * @return the converted value, or {@code null} for a blank or absent cell
     * @throws PxlCellCodecException  if a formula cell is read via the streaming reader, a formula cell is encountered without an evaluator, the column class is not a supported cell value type, or a delegated codec cannot decode the cell
     * @throws PxlReflectionException if a delegated codec fails with a reflection error
     */
    public static Object parseDataValueFromCell(final Cell cell,
                                                final PxlImportColumnMeta columnMeta,
                                                final FormulaEvaluator formulaEvaluator)
            throws PxlCellCodecException, PxlReflectionException {

        if (Objects.isNull(cell) || Objects.isNull(columnMeta)) {
            return null;
        }

        if (PxlCellUtils.isBlankCell(cell)) {
            return null;
        }

        final Field columnField = columnMeta.getColumnField();
        final Class<?> columnClass = columnMeta.getColumnClass();

        if (Objects.isNull(columnField) || Objects.isNull(columnClass)) {
            return null;
        }

        if (CellType.FORMULA.equals(cell.getCellType())) {
            if (cell instanceof StreamingCell) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_STREAMING_FORMULA_UNSUPPORTED));
            }
            if (Objects.isNull(formulaEvaluator)) {
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cell.getCellType().toString())));
            }

            formulaEvaluator.evaluateInCell(cell);
        }

        Object valueObject;

        if (columnClass == String.class) {
            valueObject = PxlStringCodec.parseStringValue(cell, columnMeta);
        } else if (columnClass == byte.class) {
            valueObject = PxlPrimitiveByteCodec.parsePrimitiveByteValue(cell, columnMeta);
        } else if (columnClass == Byte.class) {
            valueObject = PxlByteCodec.parseByteValue(cell, columnMeta);
        } else if (columnClass == short.class) {
            valueObject = PxlPrimitiveShortCodec.parsePrimitiveShortValue(cell, columnMeta);
        } else if (columnClass == Short.class) {
            valueObject = PxlShortCodec.parseShortValue(cell, columnMeta);
        } else if (columnClass == int.class) {
            valueObject = PxlPrimitiveIntCodec.parsePrimitiveIntValue(cell, columnMeta);
        } else if (columnClass == Integer.class) {
            valueObject = PxlIntegerCodec.parseIntegerValue(cell, columnMeta);
        } else if (columnClass == long.class) {
            valueObject = PxlPrimitiveLongCodec.parsePrimitiveLongValue(cell, columnMeta);
        } else if (columnClass == Long.class) {
            valueObject = PxlLongCodec.parseLongValue(cell, columnMeta);
        } else if (columnClass == double.class) {
            valueObject = PxlPrimitiveDoubleCodec.parsePrimitiveDoubleValue(cell, columnMeta);
        } else if (columnClass == Double.class) {
            valueObject = PxlDoubleCodec.parseDoubleValue(cell, columnMeta);
        } else if (columnClass == float.class) {
            valueObject = PxlPrimitiveFloatCodec.parsePrimitiveFloatValue(cell, columnMeta);
        } else if (columnClass == Float.class) {
            valueObject = PxlFloatCodec.parseFloatValue(cell, columnMeta);
        } else if (columnClass == char.class) {
            valueObject = PxlPrimitiveCharCodec.parsePrimitiveCharValue(cell, columnMeta);
        } else if (columnClass == Character.class) {
            valueObject = PxlCharacterCodec.parseCharacterValue(cell, columnMeta);
        } else if (columnClass == boolean.class || columnClass == Boolean.class) {
            valueObject = PxlBooleanCodec.parseBooleanValue(cell, columnMeta);
        } else if (columnClass == BigInteger.class) {
            valueObject = PxlBigIntegerCodec.parseBigIntegerValue(cell, columnMeta);
        } else if (columnClass == BigDecimal.class) {
            valueObject = PxlBigDecimalCodec.parseBigDecimalValue(cell, columnMeta);
        } else if (columnClass == Date.class) {
            valueObject = PxlJavaDateCodec.parseJavaDateValue(cell, columnMeta);
        } else if (columnClass == LocalDate.class) {
            valueObject = PxlLocalDateCodec.parseLocalDateValue(cell, columnMeta);
        } else if (columnClass == LocalTime.class) {
            valueObject = PxlLocalTimeCodec.parseLocalTimeValue(cell, columnMeta);
        } else if (columnClass == LocalDateTime.class) {
            valueObject = PxlLocalDateTimeCodec.parseLocalDateTimeValue(cell, columnMeta);
        } else if (columnClass == ZonedDateTime.class) {
            valueObject = PxlZonedDateTimeCodec.parseZonedDateTimeValue(cell, columnMeta);
        } else if (columnClass == OffsetTime.class) {
            valueObject = PxlOffsetTimeCodec.parseOffsetTimeValue(cell, columnMeta);
        } else if (columnClass == OffsetDateTime.class) {
            valueObject = PxlOffsetDateTimeCodec.parseOffsetDateTimeValue(cell, columnMeta);
        } else if (columnClass == Duration.class) {
            valueObject = PxlDurationCodec.parseDurationValue(cell, columnMeta);
        } else if (columnClass == Period.class) {
            valueObject = PxlPeriodCodec.parsePeriodValue(cell, columnMeta);
        } else if (columnClass == UUID.class) {
            valueObject = PxlUuidCodec.parseUuidValue(cell, columnMeta);
        } else if (columnClass.isEnum()) {
            valueObject = PxlEnumCodec.parseEnumValue(cell, columnMeta);
        } else if (PxlClassSupport.isCollectionClass(columnClass)) {
            valueObject = PxlCollectionCodec.parseCollectionValue(cell, columnMeta);
        } else if (columnMeta.isImportCustomConvertable()) {
            valueObject = PxlObjectCodec.parseObjectValue(cell, columnMeta);
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_COLUMN_TYPE_UNSUPPORTED, String.valueOf(columnClass.getSimpleName())));
        }

        return valueObject;
    }

    /**
     * Converts a raw string (typically a CSV field) to the column's declared type. Returns {@code null} when
     * {@code columnMeta} or the resolved field/class is {@code null}. Dispatches on {@code columnMeta.getColumnClass()}
     * to the matching per-type codec, which applies its own trimming and blank-to-{@code null} handling.
     *
     * @param s          the source string
     * @param columnMeta resolved import metadata for the target column (may be {@code null})
     * @return the converted value, or {@code null}
     * @throws PxlCellCodecException  if the column class is not a supported cell value type or a delegated codec cannot convert the string
     * @throws PxlReflectionException if a delegated codec fails with a reflection error
     */
    public static Object parseDataValueFromString(final String s,
                                                  final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException, PxlReflectionException {

        if (Objects.isNull(columnMeta)) {
            return null;
        }

        final Field columnField = columnMeta.getColumnField();
        final Class<?> columnClass = columnMeta.getColumnClass();

        if (Objects.isNull(columnField) || Objects.isNull(columnClass)) {
            return null;
        }

        Object valueObject;

        if (columnClass == String.class) {
            valueObject = PxlStringCodec.parseStringValue(s, columnMeta);
        } else if (columnClass == byte.class) {
            valueObject = PxlPrimitiveByteCodec.parsePrimitiveByteValue(s, columnMeta);
        } else if (columnClass == Byte.class) {
            valueObject = PxlByteCodec.parseByteValue(s, columnMeta);
        } else if (columnClass == short.class) {
            valueObject = PxlPrimitiveShortCodec.parsePrimitiveShortValue(s, columnMeta);
        } else if (columnClass == Short.class) {
            valueObject = PxlShortCodec.parseShortValue(s, columnMeta);
        } else if (columnClass == int.class) {
            valueObject = PxlPrimitiveIntCodec.parsePrimitiveIntValue(s, columnMeta);
        } else if (columnClass == Integer.class) {
            valueObject = PxlIntegerCodec.parseIntegerValue(s, columnMeta);
        } else if (columnClass == long.class) {
            valueObject = PxlPrimitiveLongCodec.parsePrimitiveLongValue(s, columnMeta);
        } else if (columnClass == Long.class) {
            valueObject = PxlLongCodec.parseLongValue(s, columnMeta);
        } else if (columnClass == double.class) {
            valueObject = PxlPrimitiveDoubleCodec.parsePrimitiveDoubleValue(s, columnMeta);
        } else if (columnClass == Double.class) {
            valueObject = PxlDoubleCodec.parseDoubleValue(s, columnMeta);
        } else if (columnClass == float.class) {
            valueObject = PxlPrimitiveFloatCodec.parsePrimitiveFloatValue(s, columnMeta);
        } else if (columnClass == Float.class) {
            valueObject = PxlFloatCodec.parseFloatValue(s, columnMeta);
        } else if (columnClass == char.class) {
            valueObject = PxlPrimitiveCharCodec.parsePrimitiveCharValue(s, columnMeta);
        } else if (columnClass == Character.class) {
            valueObject = PxlCharacterCodec.parseCharacterValue(s, columnMeta);
        } else if (columnClass == boolean.class || columnClass == Boolean.class) {
            valueObject = PxlBooleanCodec.parseBooleanValue(s, columnMeta);
        } else if (columnClass == BigInteger.class) {
            valueObject = PxlBigIntegerCodec.parseBigIntegerValue(s, columnMeta);
        } else if (columnClass == BigDecimal.class) {
            valueObject = PxlBigDecimalCodec.parseBigDecimalValue(s, columnMeta);
        } else if (columnClass == Date.class) {
            valueObject = PxlJavaDateCodec.parseJavaDateValue(s, columnMeta);
        } else if (columnClass == LocalDate.class) {
            valueObject = PxlLocalDateCodec.parseLocalDateValue(s, columnMeta);
        } else if (columnClass == LocalTime.class) {
            valueObject = PxlLocalTimeCodec.parseLocalTimeValue(s, columnMeta);
        } else if (columnClass == LocalDateTime.class) {
            valueObject = PxlLocalDateTimeCodec.parseLocalDateTimeValue(s, columnMeta);
        } else if (columnClass == ZonedDateTime.class) {
            valueObject = PxlZonedDateTimeCodec.parseZonedDateTimeValue(s, columnMeta);
        } else if (columnClass == OffsetTime.class) {
            valueObject = PxlOffsetTimeCodec.parseOffsetTimeValue(s, columnMeta);
        } else if (columnClass == OffsetDateTime.class) {
            valueObject = PxlOffsetDateTimeCodec.parseOffsetDateTimeValue(s, columnMeta);
        } else if (columnClass == Duration.class) {
            valueObject = PxlDurationCodec.parseDurationValue(s, columnMeta);
        } else if (columnClass == Period.class) {
            valueObject = PxlPeriodCodec.parsePeriodValue(s, columnMeta);
        } else if (columnClass == UUID.class) {
            valueObject = PxlUuidCodec.parseUuidValue(s, columnMeta);
        } else if (columnClass.isEnum()) {
            valueObject = PxlEnumCodec.parseEnumValue(s, columnMeta);
        } else if (PxlClassSupport.isCollectionClass(columnClass)) {
            valueObject = PxlCollectionCodec.parseCollectionValue(s, columnMeta);
        } else if (columnMeta.isImportCustomConvertable()) {
            valueObject = PxlObjectCodec.parseObjectValue(s, columnMeta);
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_COLUMN_TYPE_UNSUPPORTED, String.valueOf(columnClass.getSimpleName())));
        }

        return valueObject;
    }

    /**
     * Writes a value into a cell according to the column's declared type and returns the string form of what was
     * written. A {@code null} cell computes the string without writing anything, which is what
     * {@link #buildDataString} uses to render a value for a format that has no cells. Answers {@code null} when
     * {@code columnMeta} or the resolved column class is {@code null}. A {@code null} value (or a blank
     * {@link String}) yields the column's configured export-null string. Otherwise dispatches on
     * {@code columnMeta.getColumnClass()} to the matching per-type codec.
     *
     * @param cell       the target cell, or {@code null} to only compute the string
     * @param object     the source value
     * @param columnMeta resolved export metadata for the target column (may be {@code null})
     * @return the string form of the written value, or {@code null} when there is no metadata to write by
     * @throws PxlCellCodecException  if the column class is not a supported cell value type or a delegated codec cannot encode the value
     * @throws PxlArgumentException   if a delegated codec is given an invalid converter, styler, or unsupported target
     * @throws PxlReflectionException if a delegated codec fails with a reflection error
     */
    public static String buildDataCell(@Nullable final Cell cell,
                                       final Object object,
                                       final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException, PxlArgumentException, PxlReflectionException {

        if (Objects.isNull(columnMeta)) {
            return null;
        }

        if (Objects.isNull(object) || ((object instanceof String) && StringUtils.isBlank((String) object))) {
            final String exportNullString = columnMeta.getExportNullString();
            Optional.ofNullable(cell).ifPresent(targetCell -> targetCell.setCellValue(exportNullString));
            return exportNullString;
        }

        final Class<?> columnClass = columnMeta.getColumnClass();

        if (Objects.isNull(columnClass)) {
            return null;
        }

        if (columnClass == String.class) {
            return PxlStringCodec.buildStringCell(cell, object, columnMeta);
        } else if (columnClass == byte.class) {
            return PxlPrimitiveByteCodec.buildPrimitiveByteCell(cell, object, columnMeta);
        } else if (columnClass == Byte.class) {
            return PxlByteCodec.buildByteCell(cell, object, columnMeta);
        } else if (columnClass == short.class) {
            return PxlPrimitiveShortCodec.buildPrimitiveShortCell(cell, object, columnMeta);
        } else if (columnClass == Short.class) {
            return PxlShortCodec.buildShortCell(cell, object, columnMeta);
        } else if (columnClass == int.class) {
            return PxlPrimitiveIntCodec.buildPrimitiveIntCell(cell, object, columnMeta);
        } else if (columnClass == Integer.class) {
            return PxlIntegerCodec.buildIntegerCell(cell, object, columnMeta);
        } else if (columnClass == long.class) {
            return PxlPrimitiveLongCodec.buildPrimitiveLongCell(cell, object, columnMeta);
        } else if (columnClass == Long.class) {
            return PxlLongCodec.buildLongCell(cell, object, columnMeta);
        } else if (columnClass == double.class) {
            return PxlPrimitiveDoubleCodec.buildPrimitiveDoubleCell(cell, object, columnMeta);
        } else if (columnClass == Double.class) {
            return PxlDoubleCodec.buildDoubleCell(cell, object, columnMeta);
        } else if (columnClass == float.class) {
            return PxlPrimitiveFloatCodec.buildPrimitiveFloatCell(cell, object, columnMeta);
        } else if (columnClass == Float.class) {
            return PxlFloatCodec.buildFloatCell(cell, object, columnMeta);
        } else if (columnClass == char.class) {
            return PxlPrimitiveCharCodec.buildPrimitiveCharCell(cell, object, columnMeta);
        } else if (columnClass == Character.class) {
            return PxlCharacterCodec.buildCharacterCell(cell, object, columnMeta);
        } else if (columnClass == boolean.class || columnClass == Boolean.class) {
            return PxlBooleanCodec.buildBooleanCell(cell, object, columnMeta);
        } else if (columnClass == BigInteger.class) {
            return PxlBigIntegerCodec.buildBigIntegerCell(cell, object, columnMeta);
        } else if (columnClass == BigDecimal.class) {
            return PxlBigDecimalCodec.buildBigDecimalCell(cell, object, columnMeta);
        } else if (columnClass == Date.class) {
            return PxlJavaDateCodec.buildJavaDateCell(cell, object, columnMeta);
        } else if (columnClass == LocalDate.class) {
            return PxlLocalDateCodec.buildLocalDateCell(cell, object, columnMeta);
        } else if (columnClass == LocalTime.class) {
            return PxlLocalTimeCodec.buildLocalTimeCell(cell, object, columnMeta);
        } else if (columnClass == LocalDateTime.class) {
            return PxlLocalDateTimeCodec.buildLocalDateTimeCell(cell, object, columnMeta);
        } else if (columnClass == ZonedDateTime.class) {
            return PxlZonedDateTimeCodec.buildZonedDateTimeCell(cell, object, columnMeta);
        } else if (columnClass == OffsetTime.class) {
            return PxlOffsetTimeCodec.buildOffsetTimeCell(cell, object, columnMeta);
        } else if (columnClass == OffsetDateTime.class) {
            return PxlOffsetDateTimeCodec.buildOffsetDateTimeCell(cell, object, columnMeta);
        } else if (columnClass == Duration.class) {
            return PxlDurationCodec.buildDurationCell(cell, object, columnMeta);
        } else if (columnClass == Period.class) {
            return PxlPeriodCodec.buildPeriodCell(cell, object, columnMeta);
        } else if (columnClass == UUID.class) {
            return PxlUuidCodec.buildUuidCell(cell, object, columnMeta);
        } else if (columnClass.isEnum()) {
            return PxlEnumCodec.buildEnumCell(cell, object, columnMeta);
        } else if (PxlClassSupport.isCollectionClass(columnClass)) {
            return PxlCollectionCodec.buildCollectionCell(cell, object, columnMeta);
        } else if (columnMeta.isExportCustomConvertable()) {
            return PxlObjectCodec.buildObjectCell(cell, object, columnMeta);
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_COLUMN_TYPE_UNSUPPORTED, String.valueOf(columnClass.getSimpleName())));
        }
    }

    /**
     * Renders a value as the string an export would write, without a cell to write it into. This is
     * {@link #buildDataCell} with a {@code null} cell: the same dispatch and the same per-column settings
     * (pattern, masking, trim, export-null string, collection separator, custom converter) decide the result,
     * so a cell-less format renders a value exactly as the cell-based one would.
     *
     * @param object     the source value
     * @param columnMeta resolved export metadata for the target column (may be {@code null})
     * @return the string form of the value, or {@code null} when there is no metadata to render by
     * @throws PxlCellCodecException  if the column class is not a supported cell value type or a delegated codec cannot encode the value
     * @throws PxlArgumentException   if a delegated codec is given an invalid converter, styler, or unsupported target
     * @throws PxlReflectionException if a delegated codec fails with a reflection error
     */
    public static String buildDataString(final Object object,
                                         final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException, PxlArgumentException, PxlReflectionException {

        return buildDataCell(null, object, columnMeta);
    }

}
