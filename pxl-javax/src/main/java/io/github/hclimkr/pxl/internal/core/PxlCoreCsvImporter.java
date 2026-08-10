package io.github.hclimkr.pxl.internal.core;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.exception.*;
import io.github.hclimkr.pxl.internal.codec.PxlCellResolver;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlImportColumnMeta;
import io.github.hclimkr.pxl.internal.meta.PxlImportSheetMeta;
import io.github.hclimkr.pxl.internal.meta.PxlImportWorkbookMeta;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.internal.support.PxlReflectionSupport;
import io.github.hclimkr.pxl.option.PxlImportSheetOption;
import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;
import io.github.hclimkr.pxl.type.PxlFileFormat;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.io.ByteOrderMark;
import org.apache.commons.io.input.BOMInputStream;
import org.apache.commons.io.input.CloseShieldInputStream;
import org.apache.commons.lang3.StringUtils;

import javax.validation.Validator;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.*;

/**
 * CSV import routine
 */
public final class PxlCoreCsvImporter extends PxlAbstractImporter {

    /**
     * Prevents instantiation.
     */
    private PxlCoreCsvImporter() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses a set of CSV streams into a new instance of the given workbook class. (import)
     * Each CSV is treated as one sheet; a stream is matched to a sheet field by name, parsed,
     * and bound to that field. The result is optionally validated.
     *
     * @param workbookName   optional name assigned to the {@code @PxlWorkbookName} field (may be {@code null})
     * @param csvNames       the sheet names, one per CSV stream; must be the same size as {@code csvStreams}
     * @param csvStreams     the CSV input streams, one per sheet
     * @param workbookClass  the workbook class to instantiate and populate
     * @param workbookOption optional runtime workbook override (may be {@code null})
     * @param validator      optional bean validator applied when data validation is enabled (may be {@code null})
     * @return the populated workbook object
     * @throws PxlNullPointerException if {@code csvNames}, {@code csvStreams}, or {@code workbookClass} is {@code null}
     * @throws PxlArgumentException    if {@code csvNames} or {@code csvStreams} is empty
     * @throws PxlDataException        if the name/stream counts differ, a limit is exceeded, or a required sheet is missing
     * @throws PxlReflectionException  if instantiating a class or reading/writing a field fails
     * @throws PxlValidationException  if a bean-validation constraint on a row object is violated
     * @throws PxlCellCodecException   if a cell value cannot be decoded
     * @throws PxlIOException          if reading the CSV fails
     * @throws PxlI18nException        if the content i18n bundle cannot be found for the configured base name and locale
     */
    public static Object parseCsv(@Nullable final String workbookName,
                                  final List<String> csvNames,
                                  final List<InputStream> csvStreams,
                                  final Class<?> workbookClass,
                                  @Nullable final PxlImportWorkbookOption workbookOption,
                                  @Nullable final Validator validator)
            throws PxlNullPointerException, PxlArgumentException, PxlDataException, PxlReflectionException, PxlValidationException, PxlCellCodecException, PxlIOException, PxlI18nException {

        PxlAssertSupport.notEmpty(csvNames, "csvNames");
        PxlAssertSupport.notEmpty(csvStreams, "csvStreams");
        PxlAssertSupport.notNull(workbookClass, "workbookClass");

        if (PxlCollectionUtils.size(csvNames) != PxlCollectionUtils.size(csvStreams)) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_CSV_FILE_NAME_COUNT_MISMATCH));
        }

        final Object workbookObject = PxlReflectionSupport.newClassInstance(workbookClass);

        // Set the name on the sheet-name field.
        if (Objects.nonNull(workbookName)) {
            setWorkbookNameToWorkbookObject(workbookObject, workbookName);
        }

        final PxlImportWorkbookMeta workbookMeta = PxlImportWorkbookMeta.makeImportWorkbookMeta(workbookClass, workbookOption);
        workbookMeta.setImportFileFormat(PxlFileFormat.CSV);
        final List<PxlImportSheetOption> sheetOptions = workbookMeta.getImportSheetOptions();

        final List<PxlImportSheetMeta> sheetMetas = PxlImportSheetMeta.makeImportSheetMetas(workbookClass, workbookMeta, sheetOptions);
        workbookMeta.addImportSheetMetas(sheetMetas);

        readSheetsFromCsv(csvNames, csvStreams, sheetMetas);

        for (final PxlImportSheetMeta sheetMeta : sheetMetas) {
            // Reference the sheet that uses the given name.
            final int importSheetIndex = sheetMeta.getActualImportSheetIndex();
            if (importSheetIndex < 0) {
                continue;
            }

            final InputStream csvStream = PxlCollectionUtils.get(csvStreams, importSheetIndex);

            final Collection<Object> rowObjects = parseSheet(csvStream, sheetMeta, validator);

            if (Objects.nonNull(rowObjects)) {
                final Field sheetField = sheetMeta.getSheetField();

                PxlReflectionSupport.setFieldValue(sheetField, workbookObject, rowObjects);
            }
        }

        //PxlUtils.closeWorkbook(workbook);

        final boolean importDataValidation = workbookMeta.isImportDataValidation();
        if (importDataValidation && Objects.nonNull(validator)) {
            validateDataConstraint(validator, workbookObject, null, null);
        }

        return workbookObject;
    }

    /**
     * Parses a single CSV stream into a collection of row objects. (import)
     * A wildcard sheet option (if present) is applied, and uniqueness/validation are performed.
     *
     * @param csvName            the sheet name for the CSV
     * @param csvStream          the CSV input stream to read
     * @param rowCollectionClass the collection type instantiated to hold the row objects (e.g. {@link List}, {@link Set})
     * @param rowClass           the row class instantiated for each data row
     * @param workbookOption     optional runtime workbook override (may be {@code null})
     * @param validator          optional bean validator applied when data validation is enabled (may be {@code null})
     * @return the collection of populated row objects
     * @throws PxlNullPointerException if {@code csvName}, {@code csvStream}, {@code rowCollectionClass}, or {@code rowClass} is {@code null}
     * @throws PxlDataException        if a limit is exceeded or a required sheet is missing
     * @throws PxlArgumentException    if a configuration value is invalid
     * @throws PxlReflectionException  if instantiating a class or reading/writing a field fails
     * @throws PxlCellCodecException   if a cell value cannot be decoded
     * @throws PxlIOException          if reading the CSV fails
     * @throws PxlValidationException  if a bean-validation constraint on a row object is violated
     * @throws PxlI18nException        if the content i18n bundle cannot be found for the configured base name and locale
     */
    public static Object parseCsv(final String csvName,
                                  final InputStream csvStream,
                                  final Class<?> rowCollectionClass,
                                  final Class<?> rowClass,
                                  @Nullable final PxlImportWorkbookOption workbookOption,
                                  @Nullable final Validator validator)
            throws PxlNullPointerException, PxlDataException, PxlArgumentException, PxlReflectionException, PxlCellCodecException, PxlIOException, PxlValidationException, PxlI18nException {

        PxlAssertSupport.notNull(csvName, "csvName");
        PxlAssertSupport.notNull(csvStream, "csvStream");
        PxlAssertSupport.notNull(rowCollectionClass, "rowCollectionClass");
        PxlAssertSupport.notNull(rowClass, "rowClass");

        final PxlImportWorkbookMeta workbookMeta = PxlImportWorkbookMeta.makeImportWorkbookMeta(null, workbookOption);
        workbookMeta.setImportFileFormat(PxlFileFormat.CSV);

        final PxlImportSheetOption sheetOption = Optional.ofNullable(workbookMeta.getImportSheetOptions())
                .flatMap(options -> options.stream()
                        .filter(o -> StringUtils.equals(o.getFieldName(), PxlConstants.SHEET_FIELD_NAME_WILD_CARD))
                        .findFirst())
                .orElse(null);

        final PxlImportSheetMeta sheetMeta = PxlImportSheetMeta.makeImportSheetMeta(Collections.singletonList(csvName), rowCollectionClass, rowClass, workbookMeta, sheetOption);
        workbookMeta.addImportSheetMeta(sheetMeta);

        readSheetFromCsv(sheetMeta);

        final Collection<Object> rowObjects = parseSheet(csvStream, sheetMeta, validator);

        return rowObjects;
    }

    /**
     * Matches each sheet meta to a CSV stream by name and records its stream index. (import)
     * Names are compared after whitespace removal, ignoring case; each CSV may match at most one sheet.
     *
     * @param csvNames   the CSV/sheet names; must be the same size as {@code csvStreams}
     * @param csvStreams the CSV input streams
     * @param sheetMetas the sheet metas to resolve; their actual index/name are set on match
     * @throws PxlDataException if the name/stream counts differ, the sheet-count limit is exceeded,
     *                          a sheet or CSV matches ambiguously, or a required sheet is missing
     */
    private static void readSheetsFromCsv(final List<String> csvNames,
                                          final List<InputStream> csvStreams,
                                          final List<PxlImportSheetMeta> sheetMetas)
            throws PxlDataException {

        if (PxlCollectionUtils.size(csvNames) != PxlCollectionUtils.size(csvStreams)) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_CSV_FILE_NAME_COUNT_MISMATCH));
        }

        final int numOfSheets = PxlCollectionUtils.size(csvNames);
        if (numOfSheets > PxlConstants.IMPORT_MAX_NUMBER_OF_CSV_SHEETS) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_SHEET_COUNT_EXCEEDED, String.valueOf(PxlConstants.IMPORT_MAX_NUMBER_OF_CSV_SHEETS)));
        }

        final Set<Integer> claimedSheetStreamIndexes = new HashSet<>();

        for (final PxlImportSheetMeta sheetMeta : sheetMetas) {
            if (!sheetMeta.isImportEnabled()) {
                continue;
            }

            final String importSheetName = sheetMeta.getActualImportSheetName();
            final List<String> sheetNames = StringUtils.isBlank(importSheetName) ?
                    sheetMeta.getCandidateSheetNames() :
                    Collections.singletonList(importSheetName);

            for (int csvStreamIndex = 0; csvStreamIndex < numOfSheets; csvStreamIndex++) {
                final InputStream csvStream = PxlCollectionUtils.get(csvStreams, csvStreamIndex);
                if (Objects.isNull(csvStream)) {
                    continue;
                }

                final String csvName = StringUtils.deleteWhitespace(PxlCollectionUtils.get(csvNames, csvStreamIndex));
                if (StringUtils.isBlank(csvName)) {
                    continue;
                }

                if (matchesSheetName(sheetNames, csvName)) {
                    if (sheetMeta.getActualImportSheetIndex() >= 0) {
                        throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_SHEET_DUPLICATE, sheetNames));
                    }

                    if (!claimedSheetStreamIndexes.add(csvStreamIndex)) {
                        throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_CSV_MULTIPLE_SHEET_MATCH, csvName));
                    }

                    sheetMeta.setActualImportSheetIndex(csvStreamIndex);
                    sheetMeta.setActualImportSheetName(csvName);
                    // break;
                }
            }

            if ((sheetMeta.isRequired()) && (sheetMeta.getActualImportSheetIndex() < 0)) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_SHEET_NOT_FOUND, sheetNames));
            }
        }
    }

    /**
     * Assigns the single-stream CSV sheet meta to stream index 0 using its first candidate name. (import)
     *
     * @param sheetMeta the sheet meta to resolve; its actual index/name are set when import is enabled
     */
    private static void readSheetFromCsv(final PxlImportSheetMeta sheetMeta) {

        if (!sheetMeta.isImportEnabled()) {
            return;
        }

        sheetMeta.setActualImportSheetIndex(0);
        sheetMeta.setActualImportSheetName(PxlCollectionUtils.get(sheetMeta.getCandidateSheetNames(), 0));
    }

    /**
     * Reads one CSV stream into row objects. (import)
     * Takes the charset and the delimiter from the sheet meta, which resolved both for this sheet alone -
     * a CSV workbook is one file per sheet, so its sheets may differ in either. Strips the byte-order mark
     * for BOM-carrying charsets, parses all records with that delimiter, computes the 0-based header/data
     * row bounds from the 1-based meta values, resolves column indexes from the header row, populates the
     * {@code @PxlRowIndex} field, skips empty rows, optionally validates each row, and finally checks
     * column uniqueness.
     *
     * @param csvStream the CSV input stream to read
     * @param sheetMeta the resolved sheet meta, supplying this sheet's charset and delimiter
     * @param validator optional bean validator applied when data validation is enabled (may be {@code null})
     * @return the collection of populated row objects, or {@code null} when the sheet is disabled or has no columns
     * @throws PxlNullPointerException if {@code csvStream} or {@code sheetMeta} is {@code null}
     * @throws PxlReflectionException  if instantiating a class or reading/writing a field fails
     * @throws PxlArgumentException    if this sheet's {@code importCsvCharset} names no supported charset, its {@code importCsvDelimiter} cannot be a delimiter, or the {@code @PxlRowIndex} field type is unsupported
     * @throws PxlIOException          if the CSV cannot be read
     * @throws PxlDataException        if a limit is exceeded
     * @throws PxlCellCodecException   if a cell value cannot be decoded
     * @throws PxlValidationException  if a bean-validation constraint on a row object is violated
     */
    private static Collection<Object> parseSheet(final InputStream csvStream,
                                                 final PxlImportSheetMeta sheetMeta,
                                                 @Nullable final Validator validator)
            throws PxlNullPointerException, PxlReflectionException, PxlArgumentException, PxlIOException, PxlDataException, PxlCellCodecException, PxlValidationException {

        PxlAssertSupport.notNull(csvStream, "csvStream");
        PxlAssertSupport.notNull(sheetMeta, "sheetMeta");

        if (!sheetMeta.isImportEnabled()) {
            return null;
        }

        final PxlImportWorkbookMeta workbookMeta = sheetMeta.getWorkbookMeta();

        final Class<?> rowCollectionClass = sheetMeta.getRowCollectionClass();
        final Class<?> rowClass = sheetMeta.getRowClass();

        // Retrieve the column information.
        final List<PxlImportColumnMeta> columnMetas = PxlImportColumnMeta.makeImportColumnMetas(sheetMeta);
        sheetMeta.addImportColumnMetas(columnMetas);

        // Do nothing if there are no columns.
        if (PxlCollectionUtils.isEmpty(columnMetas)) {
            return null;
        }

//        // Reference the sheet that uses the given name.
//        final int importSheetIndex = sheetMeta.getImportSheetIndex();
//        if (importSheetIndex < 0) {
//            return null;
//        }
//
        final String sheetName = sheetMeta.getActualImportSheetName();

        final List<CSVRecord> csvRecords;
        CSVParser csvParser = null;

        final String importCsvCharset = sheetMeta.getImportCsvCharset();
        final char importCsvDelimiter = sheetMeta.getImportCsvDelimiter();

        // Both calls below reject invalid configuration with unchecked exceptions, which the IOException-only
        // try below would not catch: they would reach the builder boundary and be flattened into a
        // PxlSystemException naming neither the attribute nor its value. Normalize them here instead.
        final CSVFormat importCsvFormat;
        try {
            importCsvFormat = PxlConstants.DEFAULT_IMPORT_CSV_FORMAT
                    .builder()
                    .setDelimiter(importCsvDelimiter)
                    .build();
        } catch (RuntimeException runtimeException) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_CSV_DELIMITER_INVALID, sheetName), runtimeException);
        }

        final Charset importCharset;
        try {
            importCharset = Charset.forName(importCsvCharset);
        } catch (RuntimeException runtimeException) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_CSV_CHARSET_INVALID, sheetName, String.valueOf(importCsvCharset)), runtimeException);
        }

        try {
            // BOM handling: strip the BOM at the byte level only for charsets whose decoder does not consume the BOM itself.
            //  - UTF-8 and endian-explicit forms (UTF-16LE/BE): the decoder passes the BOM through as U+FEFF, so it must be stripped here.
            //    (If not stripped, U+FEFF remains in the first header cell and the first column fails to match.)
            //  - "UTF-16" (auto): the decoder uses the BOM to determine and consume the endianness, so it must NOT be stripped (stripping here breaks endianness detection).
            final List<ByteOrderMark> bomsToStrip = new ArrayList<>();
            bomsToStrip.add(ByteOrderMark.UTF_8);
            if (importCharset.equals(StandardCharsets.UTF_16LE)) {
                bomsToStrip.add(ByteOrderMark.UTF_16LE);
            } else if (importCharset.equals(StandardCharsets.UTF_16BE)) {
                bomsToStrip.add(ByteOrderMark.UTF_16BE);
            }

            // Protect the caller-owned original stream from being closed when CSVParser closes its internal Reader on termination.
            final InputStream bomStrippedInputStream = BOMInputStream.builder()
                    .setInputStream(CloseShieldInputStream.wrap(csvStream))
                    .setByteOrderMarks(bomsToStrip.toArray(new ByteOrderMark[0]))
                    .get();
            csvParser = CSVParser.parse(bomStrippedInputStream, importCharset, importCsvFormat);

            csvRecords = csvParser.getRecords();
        } catch (IOException ioException) {
            throw new PxlIOException(ioException);
        } finally {
            if (Objects.nonNull(csvParser)) {
                try {
                    csvParser.close();
                } catch (IOException ignored) {
                }
            }
        }

        final int firstRowNum = 0;
        final int lastRowNum = PxlCollectionUtils.size(csvRecords);

        int actualImportHeaderRowIndex = sheetMeta.getImportHeaderRowIndex();
        if (actualImportHeaderRowIndex == PxlConstants.DEFAULT_IMPORT_HEADER_ROW_INDEX) {
            actualImportHeaderRowIndex = firstRowNum;
        } else {
            actualImportHeaderRowIndex -= 1;  // Specified as 1-based, so convert to 0-based.
            actualImportHeaderRowIndex = Math.max(actualImportHeaderRowIndex, firstRowNum);
        }

        int actualImportOriginDataRowIndex = sheetMeta.getImportFirstDataRowIndex();
        if (actualImportOriginDataRowIndex == PxlConstants.DEFAULT_IMPORT_FIRST_DATA_ROW_INDEX) {
            actualImportOriginDataRowIndex = actualImportHeaderRowIndex + 1;
        } else {
            actualImportOriginDataRowIndex -= 1;  // Specified as 1-based, so convert to 0-based.
            actualImportOriginDataRowIndex = Math.max(actualImportOriginDataRowIndex, actualImportHeaderRowIndex + 1);
        }

        int actualImportBoundDataRowIndex = sheetMeta.getImportLastDataRowIndex();
        if (actualImportBoundDataRowIndex == PxlConstants.DEFAULT_IMPORT_LAST_DATA_ROW_INDEX) {
            actualImportBoundDataRowIndex = lastRowNum;
        } else {
            actualImportBoundDataRowIndex -= 1;  // Specified as 1-based, so convert to 0-based.
            actualImportBoundDataRowIndex += 1;  // Add 1 to make it exclusive.
            actualImportBoundDataRowIndex = Math.min(actualImportBoundDataRowIndex, lastRowNum);
        }

        final int maxNumOfRows = workbookMeta.getImportFileFormat().getMaxImportRows();
        if (actualImportBoundDataRowIndex - actualImportOriginDataRowIndex > maxNumOfRows) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_SHEET_ROW_COUNT_EXCEEDED, sheetName, String.valueOf(maxNumOfRows)));
        }

        sheetMeta.setActualImportHeaderRowIndex(actualImportHeaderRowIndex);
        sheetMeta.setActualImportOriginDataRowIndex(actualImportOriginDataRowIndex);
        sheetMeta.setActualImportBoundDataRowIndex(actualImportBoundDataRowIndex);

        // Read the first row (header row).
        readHeaderRowFromSheet(sheetName, csvRecords, sheetMeta, actualImportHeaderRowIndex);

        final boolean noImportColumn = columnMetas.stream().allMatch(c -> c.getActualImportColumnIndex() < 0);
        if (noImportColumn) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_SHEET_NO_HEADER_COLUMN, sheetName, String.valueOf(actualImportHeaderRowIndex + 1)));
        }

        // Obtain the row index field.
        final Field rowIndexField = getRowIndexField(rowClass);
        Class<?> rowIndexClass = null;
        if (Objects.nonNull(rowIndexField)) {
            rowIndexClass = rowIndexField.getType();
        }

        final boolean importDataValidation = workbookMeta.isImportDataValidation();

        final Collection<Object> rowObjects = (Collection<Object>) PxlReflectionSupport.newClassInstance(rowCollectionClass);

        // Read each data row.
        for (int rowIndex = actualImportOriginDataRowIndex; rowIndex < actualImportBoundDataRowIndex; rowIndex++) {

            final Object rowObject = PxlReflectionSupport.newClassInstance(rowClass);

            if (Objects.nonNull(rowIndexField) && Objects.nonNull(rowIndexClass)) {
                // Expose the 0-based record index as a 1-based row number.
                final int oneBasedRowIndex = rowIndex + 1;
                try {
                    if (rowIndexClass == Long.class || rowIndexClass == long.class) {
                        PxlReflectionSupport.setFieldValue(rowIndexField, rowObject, Long.valueOf(oneBasedRowIndex));
                    } else if (rowIndexClass == Integer.class || rowIndexClass == int.class) {
                        PxlReflectionSupport.setFieldValue(rowIndexField, rowObject, Integer.valueOf(oneBasedRowIndex));
                    } else if (rowIndexClass == Short.class || rowIndexClass == short.class) {
                        PxlReflectionSupport.setFieldValue(rowIndexField, rowObject, Short.valueOf((short) oneBasedRowIndex));
                    } else if (rowIndexClass == Byte.class || rowIndexClass == byte.class) {
                        PxlReflectionSupport.setFieldValue(rowIndexField, rowObject, Byte.valueOf((byte) oneBasedRowIndex));
                    } else {
                        PxlReflectionSupport.setFieldValue(rowIndexField, rowObject, rowIndexClass.cast(oneBasedRowIndex));
                    }
                } catch (Exception e) {
                    throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_ROW_INDEX_TYPE_UNSUPPORTED, rowIndexField.getName(), rowIndexField.getType().getSimpleName()), e);
                }
            }

            readDataRowFromSheet(sheetName, csvRecords, columnMetas, rowObject, rowIndex);

            // Skip the row if it is ignorable.
            if (isIgnorableRow(columnMetas, rowObject)) {
                continue;
            }

            if (importDataValidation && Objects.nonNull(validator)) {
                validateDataConstraint(validator, rowObject, sheetName, rowIndex);
            }

            rowObjects.add(rowObject);
        }

        validateDataUniqueness(rowObjects, columnMetas, sheetName);

        return rowObjects;
    }

    /**
     * Reads the header record and resolves each column's actual index by matching its header cell value
     * against the column's candidate names. (import)
     * Header cell values are compared after whitespace removal.
     *
     * @param csvName        the CSV/sheet name used in error messages
     * @param csvRecords     the parsed CSV records
     * @param sheetMeta      the resolved sheet meta whose column metas are updated with matched indexes/names
     * @param headerRowIndex the 0-based index of the header record within {@code csvRecords}
     * @throws PxlDataException if the header record is missing, a column exceeds the column limit,
     *                          a column name is duplicated, or a required column is missing
     */
    private static void readHeaderRowFromSheet(final String csvName,
                                               final List<CSVRecord> csvRecords,
                                               final PxlImportSheetMeta sheetMeta,
                                               final int headerRowIndex)
            throws PxlDataException {

        final PxlImportWorkbookMeta workbookMeta = sheetMeta.getWorkbookMeta();
        final List<PxlImportColumnMeta> columnMetas = sheetMeta.getImportColumnMetas();

        final CSVRecord csvRecord = PxlCollectionUtils.get(csvRecords, headerRowIndex);
        if (Objects.isNull(csvRecord)) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_SHEET_NO_HEADER_ROW, csvName, String.valueOf(headerRowIndex + 1)));
        }

        final int firstCellNum = 0;
        final int lastCellNum = csvRecord.size();
/*
        if (firstCellNum < 0 || lastCellNum < 0) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_SHEET_NO_HEADER_ROW, csvName, String.valueOf(headerRowIndex + 1)));
        }
*/

        int actualImportOriginDataColumnIndex = sheetMeta.getImportFirstDataColumnIndex();
        if (actualImportOriginDataColumnIndex == PxlConstants.DEFAULT_IMPORT_FIRST_DATA_COLUMN_INDEX) {
            actualImportOriginDataColumnIndex = firstCellNum;
        } else {
            actualImportOriginDataColumnIndex -= 1;  // Specified as 1-based, so convert to 0-based.
            actualImportOriginDataColumnIndex = Math.max(actualImportOriginDataColumnIndex, firstCellNum);
        }

        int actualImportBoundDataColumnIndex = sheetMeta.getImportLastDataColumnIndex();
        if (actualImportBoundDataColumnIndex == PxlConstants.DEFAULT_IMPORT_LAST_DATA_COLUMN_INDEX) {
            actualImportBoundDataColumnIndex = lastCellNum;
        } else {
            actualImportBoundDataColumnIndex -= 1;  // Specified as 1-based, so convert to 0-based.
            actualImportBoundDataColumnIndex += 1;  // Add 1 to make it exclusive.
            actualImportBoundDataColumnIndex = Math.min(actualImportBoundDataColumnIndex, lastCellNum);
        }

        final int maxNumOfColumns = workbookMeta.getImportFileFormat().getMaxImportColumns();
        if (actualImportBoundDataColumnIndex - actualImportOriginDataColumnIndex > maxNumOfColumns) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_SHEET_COLUMN_COUNT_EXCEEDED, csvName, String.valueOf(maxNumOfColumns)));
        }

        sheetMeta.setActualImportOriginDataColumnIndex(actualImportOriginDataColumnIndex);
        sheetMeta.setActualImportBoundDataColumnIndex(actualImportBoundDataColumnIndex);

        for (final PxlImportColumnMeta columnMeta : columnMetas) {
            if (!columnMeta.isImportEnabled()) {
                continue;
            }

            final List<String> candidateColumnNames = columnMeta.getCandidateColumnNames();

            for (int importColumnIndex = actualImportOriginDataColumnIndex; importColumnIndex < actualImportBoundDataColumnIndex; importColumnIndex++) {
                final String columnName = StringUtils.deleteWhitespace(PxlCollectionUtils.get(csvRecord.values(), importColumnIndex));
                if (StringUtils.isBlank(columnName)) {
                    continue;
                }

                if (candidateColumnNames.contains(columnName)) {
                    if (columnMeta.getActualImportColumnIndex() >= 0) {
                        throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_COLUMN_DUPLICATE, csvName, candidateColumnNames));
                    }

                    columnMeta.setActualImportColumnIndex(importColumnIndex);
                    columnMeta.setActualImportColumnName(columnName);
                    // break;
                }
            }

            if ((columnMeta.isRequired()) && (columnMeta.getActualImportColumnIndex() < 0)) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_COLUMN_NOT_FOUND, csvName, candidateColumnNames));
            }
        }
    }

    /**
     * Reads one CSV record and decodes each mapped field value into the corresponding field of the row object. (import)
     * Missing records and {@code null} field values are skipped.
     *
     * @param csvName     the CSV/sheet name used in error messages
     * @param csvRecords  the parsed CSV records
     * @param columnMetas the per-column import metadata
     * @param rowObject   the row object populated with decoded values
     * @param rowIndex    the 0-based index of the data record within {@code csvRecords}
     * @throws PxlCellCodecException if a field value cannot be decoded into the target field type
     */
    private static void readDataRowFromSheet(final String csvName,
                                             final List<CSVRecord> csvRecords,
                                             final List<PxlImportColumnMeta> columnMetas,
                                             final Object rowObject,
                                             final int rowIndex)
            throws PxlCellCodecException {

        final CSVRecord csvRecord = PxlCollectionUtils.get(csvRecords, rowIndex);
        if (Objects.isNull(csvRecord)) {
            return;
        }

        for (final PxlImportColumnMeta columnMeta : columnMetas) {
            if (!columnMeta.isImportEnabled()) {
                continue;
            }

            final int importColumnIndex = columnMeta.getActualImportColumnIndex();
            if (importColumnIndex < 0) {
                continue;
            }

            final String csvValue = PxlCollectionUtils.get(csvRecord.values(), importColumnIndex);
            if (Objects.isNull(csvValue)) {
                continue;
            }

            try {
                final Object valueObject = PxlCellResolver.parseDataValueFromString(csvValue, columnMeta);
                final Field columnField = columnMeta.getColumnField();

                if (Objects.nonNull(valueObject)) {
                    PxlReflectionSupport.setFieldValue(columnField, rowObject, valueObject);
                }
            } catch (Exception e) {
                throw new PxlCellCodecException(csvName, rowIndex, columnMeta.getActualImportColumnName(), importColumnIndex, e);
            }
        }
    }

}
