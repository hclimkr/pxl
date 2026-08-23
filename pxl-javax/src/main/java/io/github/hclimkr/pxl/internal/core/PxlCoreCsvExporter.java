package io.github.hclimkr.pxl.internal.core;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.exception.*;
import io.github.hclimkr.pxl.internal.codec.PxlCellResolver;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlExportColumnMeta;
import io.github.hclimkr.pxl.internal.meta.PxlExportSheetMeta;
import io.github.hclimkr.pxl.internal.meta.PxlExportWorkbookMeta;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.internal.support.PxlReflectionSupport;
import io.github.hclimkr.pxl.option.PxlExportSheetOption;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.lang3.StringUtils;

import javax.validation.Validator;
import java.io.IOException;
import java.io.Writer;
import java.lang.reflect.Field;
import java.util.*;

/**
 * Core CSV writer for export.
 *
 * <p>Writes one sheet to a {@link Writer}: the charset, the byte order mark and the destination are the builder's
 * business, and only the CSV grammar is this class's. The values themselves come from the same codec entry point
 * the Excel exporter uses, called in its cell-less form
 * ({@code PxlCellResolver.buildDataString}), so a type behaves identically in both formats.</p>
 */
public final class PxlCoreCsvExporter extends PxlAbstractExporter {

    /**
     * Not instantiable.
     */
    private PxlCoreCsvExporter() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Writes one sheet of row objects as CSV. (export)
     *
     * <p>The writer is flushed but never closed - it belongs to the caller.</p>
     *
     * @param sheetName    the sheet name, used for the sheet meta and in error messages
     * @param rowObjects   the row objects written as data records
     * @param rowClass     the row class describing the column bindings
     * @param workbookMeta the resolved export metadata for the workbook
     * @param validator    optional bean validator applied when data validation is enabled (may be {@code null})
     * @param writer       the destination the records are printed to
     * @throws PxlNullPointerException if {@code sheetName}, {@code rowObjects}, {@code rowClass}, {@code workbookMeta}, or {@code writer} is {@code null}
     * @throws PxlArgumentException    if {@code sheetName} is blank, a configuration value is invalid, or the field delimiter cannot be used
     * @throws PxlValidationException  if a bean-validation constraint on a row object is violated
     * @throws PxlDataException        if a row object is {@code null}, a limit is exceeded, or there is no data to write
     * @throws PxlCellCodecException   if a cell value cannot be encoded
     * @throws PxlReflectionException  if a column field's type cannot be resolved
     * @throws PxlIOException          if writing to the destination fails
     */
    public static void writeCsv(final String sheetName,
                                final Collection<?> rowObjects,
                                final Class<?> rowClass,
                                final PxlExportWorkbookMeta workbookMeta,
                                @Nullable final Validator validator,
                                final Writer writer)
            throws PxlNullPointerException, PxlArgumentException, PxlValidationException, PxlDataException, PxlCellCodecException, PxlReflectionException, PxlIOException {

        PxlAssertSupport.notBlank(sheetName, "sheetName");
        PxlAssertSupport.notNull(rowObjects, "rowObjects");
        PxlAssertSupport.notNull(rowClass, "rowClass");
        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");
        PxlAssertSupport.notNull(writer, "writer");

        final boolean exportDataValidation = workbookMeta.isExportDataValidation();
        if (exportDataValidation && Objects.nonNull(validator)) {
            validateBeanConstraints(validator, rowObjects, null, null);
        }

        final PxlExportSheetMeta sheetMeta = makeSheetMeta(sheetName, rowObjects.getClass(), rowClass, workbookMeta);

        if (!sheetMeta.isExportEnabled()) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_NO_DATA));
        }

        // exportIfNull has nothing to act on here: the sheet form takes the collection as an argument, which was
        // rejected above if null. It stays meaningful for the workbook form, where the collection is a field.
        if (!sheetMeta.isExportIfEmpty() && rowObjects.isEmpty()) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_NO_DATA));
        }

        final int numOfObjects = PxlCollectionUtils.size(rowObjects);

        applyExportRowIndices(sheetMeta, numOfObjects);
        final int actualExportHeaderRowIndex = sheetMeta.getActualExportHeaderRowIndex();
        final int actualExportOriginDataRowIndex = sheetMeta.getActualExportOriginDataRowIndex();
        final int actualExportBoundDataRowIndex = sheetMeta.getActualExportBoundDataRowIndex();

        final int maxNumOfRows = workbookMeta.getExportFileFormat().getMaxExportRows();
        if (actualExportBoundDataRowIndex > maxNumOfRows) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_SHEET_ROW_COUNT_EXCEEDED,
                    sheetMeta.getActualExportSheetName(), String.valueOf(maxNumOfRows)));
        }

        final List<PxlExportColumnMeta> columnMetas = PxlExportColumnMeta.makeExportColumnMetas(sheetMeta, false);
        sheetMeta.addExportColumnMetas(columnMetas);

        if (PxlCollectionUtils.isEmpty(columnMetas)) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_NO_DATA));
        }

        int rowPosition = 0;
        for (final Object rowObject : rowObjects) {
            rowPosition++;
            if (Objects.isNull(rowObject)) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_ROW_NULL,
                        sheetMeta.getActualExportSheetName(), String.valueOf(rowPosition)));
            }
        }

        final CSVPrinter csvPrinter = makeCsvPrinter(sheetMeta, writer);
        final int numOfFields = countFields(columnMetas);

        try {
            writePaddingRecords(csvPrinter, 0, actualExportHeaderRowIndex, numOfFields);
            writeHeaderRecord(csvPrinter, columnMetas, numOfFields);
            writePaddingRecords(csvPrinter, actualExportHeaderRowIndex + 1, actualExportOriginDataRowIndex, numOfFields);

            int rowIndex = actualExportOriginDataRowIndex;
            for (final Object rowObject : rowObjects) {
                if (rowIndex >= actualExportBoundDataRowIndex) {
                    break;
                }
                writeDataRecord(csvPrinter, sheetMeta, columnMetas, rowObject, rowIndex);
                rowIndex++;
            }

            csvPrinter.flush();
        } catch (IOException e) {
            throw new PxlIOException(e);
        }
    }

    /**
     * Writes one sample (template) sheet as CSV: a header record and a single record of the columns'
     * {@code exportSample} values. (export)
     *
     * <p>The writer is flushed but never closed - it belongs to the caller.</p>
     *
     * @param sheetName    the sheet name, used for the sheet meta and in error messages
     * @param rowClass     the row class describing the column bindings
     * @param workbookMeta the resolved export metadata for the workbook
     * @param writer       the destination the records are printed to
     * @throws PxlNullPointerException if {@code sheetName}, {@code rowClass}, {@code workbookMeta}, or {@code writer} is {@code null}
     * @throws PxlArgumentException    if {@code sheetName} is blank, a configuration value is invalid, or the field delimiter cannot be used
     * @throws PxlDataException        if a limit is exceeded or there is no data to write
     * @throws PxlCellCodecException   if a sample value cannot be encoded
     * @throws PxlReflectionException  if a column field's type cannot be resolved
     * @throws PxlIOException          if writing to the destination fails
     */
    public static void writeSampleCsv(final String sheetName,
                                      final Class<?> rowClass,
                                      final PxlExportWorkbookMeta workbookMeta,
                                      final Writer writer)
            throws PxlNullPointerException, PxlArgumentException, PxlDataException, PxlCellCodecException, PxlReflectionException, PxlIOException {

        PxlAssertSupport.notBlank(sheetName, "sheetName");
        PxlAssertSupport.notNull(rowClass, "rowClass");
        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");
        PxlAssertSupport.notNull(writer, "writer");

        final PxlExportSheetMeta sheetMeta = makeSheetMeta(sheetName, null, rowClass, workbookMeta);

        if (!sheetMeta.isExportSampleEnabled()) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_NO_DATA));
        }

        // A sample always carries exactly one data row, so the declared data bound plays no part.
        applyExportRowIndices(sheetMeta, 1);
        final int actualExportHeaderRowIndex = sheetMeta.getActualExportHeaderRowIndex();
        final int actualExportOriginDataRowIndex = sheetMeta.getActualExportOriginDataRowIndex();

        final List<PxlExportColumnMeta> columnMetas = PxlExportColumnMeta.makeExportColumnMetas(sheetMeta, true);
        sheetMeta.addExportColumnMetas(columnMetas);

        if (PxlCollectionUtils.isEmpty(columnMetas)) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_NO_DATA));
        }

        final CSVPrinter csvPrinter = makeCsvPrinter(sheetMeta, writer);
        final int numOfFields = countFields(columnMetas);

        try {
            writePaddingRecords(csvPrinter, 0, actualExportHeaderRowIndex, numOfFields);
            writeHeaderRecord(csvPrinter, columnMetas, numOfFields);
            writePaddingRecords(csvPrinter, actualExportHeaderRowIndex + 1, actualExportOriginDataRowIndex, numOfFields);

            writeSampleRecord(csvPrinter, sheetMeta, columnMetas);

            csvPrinter.flush();
        } catch (IOException e) {
            throw new PxlIOException(e);
        }
    }

    /**
     * Writes one empty record per row in the given range, standing in for the rows a sheet leaves blank above the
     * header or between the header and the first data row.
     *
     * <p>Each such row is printed as a full record of empty fields rather than as an empty line. PXL's own CSV
     * import ignores empty lines, so a blank line would vanish on the way back in and pull the header up - an
     * empty-field record survives because its first field is quoted.</p>
     *
     * @param csvPrinter   the printer to write to
     * @param fromRowIndex the first 0-based row index to fill, inclusive
     * @param toRowIndex   the row index to stop before, exclusive
     * @param numOfFields  the number of fields per record
     * @throws IOException if writing fails
     */
    private static void writePaddingRecords(final CSVPrinter csvPrinter,
                                            final int fromRowIndex,
                                            final int toRowIndex,
                                            final int numOfFields)
            throws IOException {

        for (int rowIndex = fromRowIndex; rowIndex < toRowIndex; rowIndex++) {
            csvPrinter.printRecord(makeEmptyRecord(numOfFields));
        }
    }

    /**
     * Writes the header record, placing each mapped column's name in the field at that column's index and
     * leaving every unmapped field empty.
     *
     * @param csvPrinter  the printer to write to
     * @param columnMetas the per-column export metadata
     * @param numOfFields the number of fields per record
     * @throws IOException if writing fails
     */
    private static void writeHeaderRecord(final CSVPrinter csvPrinter,
                                          final List<PxlExportColumnMeta> columnMetas,
                                          final int numOfFields)
            throws IOException {

        final List<String> record = makeEmptyRecord(numOfFields);
        for (final PxlExportColumnMeta columnMeta : columnMetas) {
            final int exportColumnIndex = columnMeta.getActualExportColumnIndex();
            if (exportColumnIndex < 0) {
                continue;
            }
            record.set(exportColumnIndex, columnMeta.getActualExportColumnName());
        }

        csvPrinter.printRecord(record);
    }

    /**
     * Writes one data record, reading each mapped field of the row object and encoding it into a string.
     *
     * <p>A mapped column always occupies a field, even when the codec answers {@code null}: dropping it would
     * shift every later column of that record. {@code null} values are intentionally not skipped so that
     * {@code exportNullString} can take effect.</p>
     *
     * @param csvPrinter  the printer to write to
     * @param sheetMeta   the sheet meta, supplying the sheet name for error messages
     * @param columnMetas the per-column export metadata
     * @param rowObject   the row object whose field values are written
     * @param rowIndex    the 0-based index of the record being written
     * @throws IOException           if writing fails
     * @throws PxlCellCodecException if a field value cannot be read or encoded
     */
    private static void writeDataRecord(final CSVPrinter csvPrinter,
                                        final PxlExportSheetMeta sheetMeta,
                                        final List<PxlExportColumnMeta> columnMetas,
                                        final Object rowObject,
                                        final int rowIndex)
            throws IOException, PxlCellCodecException {

        final String sheetName = sheetMeta.getActualExportSheetName();
        final List<String> record = makeEmptyRecord(countFields(columnMetas));

        for (final PxlExportColumnMeta columnMeta : columnMetas) {
            final int exportColumnIndex = columnMeta.getActualExportColumnIndex();
            if (exportColumnIndex < 0) {
                continue;
            }

            final Field columnField = columnMeta.getColumnField();

            final Object cellObject;
            try {
                cellObject = PxlReflectionSupport.getFieldValue(columnField, rowObject);
            } catch (Exception e) {
                throw new PxlCellCodecException(sheetName, rowIndex, columnMeta.getActualExportColumnName(), exportColumnIndex, e);
            }

            try {
                record.set(exportColumnIndex, PxlCellResolver.buildDataString(cellObject, columnMeta));
            } catch (Exception e) {
                throw new PxlCellCodecException(sheetName, rowIndex, columnMeta.getActualExportColumnName(), exportColumnIndex, e);
            }
        }

        csvPrinter.printRecord(record);
    }

    /**
     * Writes the single sample record, holding each column's {@code exportSample} value.
     *
     * <p>A blank sample is passed through rather than skipped, so that {@code exportNullString} can take effect
     * just as it does for a data record.</p>
     *
     * @param csvPrinter  the printer to write to
     * @param sheetMeta   the sheet meta, supplying the sheet name and the record's row index for error messages
     * @param columnMetas the per-column export metadata supplying the sample values
     * @throws IOException           if writing fails
     * @throws PxlCellCodecException if a sample value cannot be encoded
     */
    private static void writeSampleRecord(final CSVPrinter csvPrinter,
                                          final PxlExportSheetMeta sheetMeta,
                                          final List<PxlExportColumnMeta> columnMetas)
            throws IOException, PxlCellCodecException {

        final String sheetName = sheetMeta.getActualExportSheetName();
        final int rowIndex = sheetMeta.getActualExportOriginDataRowIndex();
        final List<String> record = makeEmptyRecord(countFields(columnMetas));

        for (final PxlExportColumnMeta columnMeta : columnMetas) {
            final int exportColumnIndex = columnMeta.getActualExportColumnIndex();
            if (exportColumnIndex < 0) {
                continue;
            }

            try {
                record.set(exportColumnIndex, PxlCellResolver.buildDataString(columnMeta.getExportSample(), columnMeta));
            } catch (Exception e) {
                throw new PxlCellCodecException(sheetName, rowIndex, columnMeta.getActualExportColumnName(), exportColumnIndex, e);
            }
        }

        csvPrinter.printRecord(record);
    }

    /**
     * Creates the sheet meta for an ad-hoc CSV sheet, applying the wildcard sheet option if one is registered.
     *
     * @param sheetName          the sheet name
     * @param rowCollectionClass the Collection type holding the rows, or {@code null} for a sample
     * @param rowClass           the row class
     * @param workbookMeta       the resolved export metadata for the workbook
     * @return the resolved sheet meta
     * @throws PxlNullPointerException if a required argument is null
     * @throws PxlArgumentException    if {@code sheetName} is blank or a configuration value is invalid
     * @throws PxlDataException        if {@code rowCollectionClass} is not a Collection type, or a row/column index is negative or inconsistent
     */
    private static PxlExportSheetMeta makeSheetMeta(final String sheetName,
                                                    @Nullable final Class<?> rowCollectionClass,
                                                    final Class<?> rowClass,
                                                    final PxlExportWorkbookMeta workbookMeta)
            throws PxlNullPointerException, PxlArgumentException, PxlDataException {

        final PxlExportSheetOption sheetOption = Optional.ofNullable(workbookMeta.getExportSheetOptions())
                .flatMap(options -> options.stream()
                        .filter(o -> StringUtils.equals(o.getFieldName(), PxlConstants.SHEET_FIELD_NAME_WILD_CARD))
                        .findFirst())
                .orElse(null);

        final PxlExportSheetMeta sheetMeta = PxlExportSheetMeta.makeExportSheetMeta(sheetName, rowCollectionClass, rowClass, workbookMeta, sheetOption);
        workbookMeta.addExportSheetMeta(sheetMeta);

        return sheetMeta;
    }

    /**
     * Creates the printer for this sheet, applying the resolved field delimiter.
     *
     * <p>The builder rejects an unusable delimiter before the destination is opened; this repeats the check as a
     * backstop for callers reaching the core directly, so the failure keeps naming the offending value instead of
     * surfacing as an unclassified system error.</p>
     *
     * @param sheetMeta the sheet meta supplying the delimiter
     * @param writer    the destination
     * @return the printer to write records with
     * @throws PxlArgumentException if the delimiter cannot be used
     * @throws PxlIOException       if the printer cannot be created
     */
    private static CSVPrinter makeCsvPrinter(final PxlExportSheetMeta sheetMeta,
                                             final Writer writer)
            throws PxlArgumentException, PxlIOException {

        try {
            final CSVFormat csvFormat = PxlConstants.DEFAULT_EXPORT_CSV_FORMAT
                    .builder()
                    .setDelimiter(sheetMeta.getExportCsvDelimiter())
                    .build();

            return new CSVPrinter(writer, csvFormat);
        } catch (IllegalArgumentException e) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.BUILDER_EXPORT_CSV_DELIMITER_INVALID,
                    sheetMeta.getActualExportSheetName(), String.valueOf(sheetMeta.getExportCsvDelimiter())));
        } catch (IOException e) {
            throw new PxlIOException(e);
        }
    }

    /**
     * Returns how many fields a record holds: the highest column index in use plus one, so that a first data
     * column above zero leaves the leading fields empty. Answers zero when no column is mapped.
     *
     * @param columnMetas the per-column export metadata
     * @return the number of fields per record
     */
    private static int countFields(final List<PxlExportColumnMeta> columnMetas) {

        int maxColumnIndex = -1;
        for (final PxlExportColumnMeta columnMeta : columnMetas) {
            maxColumnIndex = Math.max(maxColumnIndex, columnMeta.getActualExportColumnIndex());
        }

        return maxColumnIndex + 1;
    }

    /**
     * Creates a record of empty fields.
     *
     * @param numOfFields the number of fields
     * @return the record values
     */
    private static List<String> makeEmptyRecord(final int numOfFields) {

        final List<String> record = new ArrayList<>(numOfFields);
        for (int index = 0; index < numOfFields; index++) {
            record.add("");
        }

        return record;
    }

}
