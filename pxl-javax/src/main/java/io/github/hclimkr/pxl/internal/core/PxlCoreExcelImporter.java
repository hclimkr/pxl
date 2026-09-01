package io.github.hclimkr.pxl.internal.core;

import com.github.pjfanning.xlsx.impl.StreamingSheet;
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
import io.github.hclimkr.pxl.util.PxlCellUtils;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import io.github.hclimkr.pxl.util.PxlWorkbookUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;

import javax.validation.Validator;
import java.lang.reflect.Field;
import java.util.*;

/**
 * Excel import routine
 */
public final class PxlCoreExcelImporter extends PxlAbstractImporter {

    /**
     * Prevents instantiation.
     */
    private PxlCoreExcelImporter() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses an open POI workbook into a new instance of the given workbook class. (import)
     * Resolves the sheet metas, reads every mapped sheet, populates the corresponding sheet fields,
     * closes the workbook, and optionally validates the resulting object.
     *
     * @param workbookName  optional name assigned to the {@code @PxlWorkbookName} field (may be {@code null})
     * @param workbook      the open POI workbook to read from
     * @param workbookClass the workbook class to instantiate and populate
     * @param workbookMeta  the resolved import metadata for the workbook
     * @param sheetOptions  optional per-sheet runtime overrides (may be {@code null})
     * @param validator     optional bean validator applied when data validation is enabled (may be {@code null})
     * @return the populated workbook object
     * @throws PxlNullPointerException if {@code workbook}, {@code workbookClass}, or {@code workbookMeta} is {@code null}
     * @throws PxlReflectionException  if instantiating a class or reading/writing a field fails
     * @throws PxlDataException        if a limit is exceeded, or a required or duplicate sheet check fails
     * @throws PxlArgumentException    if stream reading conflicts with merged-region handling
     * @throws PxlCellCodecException   if a cell value cannot be decoded
     * @throws PxlValidationException  if a bean-validation constraint on a row object is violated
     */
    public static Object parseExcel(@Nullable final String workbookName,
                                    final Workbook workbook,
                                    final Class<?> workbookClass,
                                    final PxlImportWorkbookMeta workbookMeta,
                                    @Nullable final List<PxlImportSheetOption> sheetOptions,
                                    @Nullable final Validator validator)
            throws PxlNullPointerException, PxlReflectionException, PxlDataException, PxlArgumentException, PxlCellCodecException, PxlValidationException {

        PxlAssertSupport.notNull(workbook, "workbook");
        PxlAssertSupport.notNull(workbookClass, "workbookClass");
        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");

        final Object workbookObject = PxlReflectionSupport.newClassInstance(workbookClass);

        // Set the name on the sheet-name field.
        if (Objects.nonNull(workbookName)) {
            injectWorkbookName(workbookObject, workbookName);
        }

        final List<PxlImportSheetMeta> sheetMetas = PxlImportSheetMeta.makeImportSheetMetas(workbookClass, workbookMeta, sheetOptions);
        workbookMeta.addImportSheetMetas(sheetMetas);

        if (workbookMeta.isImportUsingStreamReader()) {
            final boolean anySheetImportEachCellOfMergedRegion = sheetMetas.stream()
                    .anyMatch(PxlImportSheetMeta::isImportEachCellOfMergedRegion);

            if (anySheetImportEachCellOfMergedRegion) {
                throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_STREAMING_MERGED_UNSUPPORTED));
            }
        }

        resolveAllSheetsFromWorkbook(workbook, workbookMeta, sheetMetas);

        final FormulaEvaluator formulaEvaluator = PxlWorkbookUtils.createFormulaEvaluator(workbook);

        for (final PxlImportSheetMeta sheetMeta : sheetMetas) {
            final Collection<?> rowObjects = parseSheet(workbook, sheetMeta, formulaEvaluator, validator);

            if (Objects.nonNull(rowObjects)) {
                final Field sheetField = sheetMeta.getSheetField();

                PxlReflectionSupport.setFieldValue(sheetField, workbookObject, rowObjects);
            }
        }

        PxlWorkbookUtils.closeWorkbook(workbook);

        final boolean importDataValidation = workbookMeta.isImportDataValidation();
        if (importDataValidation && Objects.nonNull(validator)) {
            validateBeanConstraints(validator, workbookObject, null, null);
        }

        return workbookObject;
    }

    /**
     * Parses a single sheet of an open POI workbook into a collection of row objects. (import)
     * The sheet is located by matching any of the candidate names, then read, and the workbook is closed.
     *
     * @param workbook            the open POI workbook to read from
     * @param candidateSheetNames the accepted sheet names; the first matching sheet is used
     * @param rowCollectionClass  the collection type instantiated to hold the row objects (e.g. {@link List}, {@link Set})
     * @param rowClass            the row class instantiated for each data row
     * @param workbookMeta        the resolved import metadata for the workbook
     * @param sheetOption         optional per-sheet runtime override (may be {@code null})
     * @param validator           optional bean validator applied when data validation is enabled (may be {@code null})
     * @return the collection of populated row objects
     * @throws PxlNullPointerException if {@code workbook}, {@code candidateSheetNames}, {@code rowCollectionClass}, {@code rowClass}, or {@code workbookMeta} is {@code null}
     * @throws PxlDataException        if a limit is exceeded, or a required or duplicate sheet check fails
     * @throws PxlArgumentException    if {@code candidateSheetNames} is empty, or stream reading conflicts with merged-region handling
     * @throws PxlCellCodecException   if a cell value cannot be decoded
     * @throws PxlReflectionException  if instantiating a class or reading/writing a field fails
     * @throws PxlValidationException  if a bean-validation constraint on a row object is violated
     */
    public static Collection<?> parseExcel(final Workbook workbook,
                                           final List<String> candidateSheetNames,
                                           final Class<?> rowCollectionClass,
                                           final Class<?> rowClass,
                                           final PxlImportWorkbookMeta workbookMeta,
                                           @Nullable final PxlImportSheetOption sheetOption,
                                           @Nullable final Validator validator)
            throws PxlNullPointerException, PxlDataException, PxlArgumentException, PxlCellCodecException, PxlReflectionException, PxlValidationException {

        PxlAssertSupport.notNull(workbook, "workbook");
        PxlAssertSupport.notEmpty(candidateSheetNames, "candidateSheetNames");
        PxlAssertSupport.notNull(rowCollectionClass, "rowCollectionClass");
        PxlAssertSupport.notNull(rowClass, "rowClass");
        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");

        final PxlImportSheetMeta sheetMeta = PxlImportSheetMeta.makeImportSheetMeta(candidateSheetNames, rowCollectionClass, rowClass, workbookMeta, sheetOption);
        workbookMeta.addImportSheetMeta(sheetMeta);

        if (workbookMeta.isImportUsingStreamReader()) {
            if (sheetMeta.isImportEachCellOfMergedRegion()) {
                throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_STREAMING_MERGED_UNSUPPORTED));
            }
        }

        resolveSheetFromWorkbook(workbook, sheetMeta);

        final FormulaEvaluator formulaEvaluator = PxlWorkbookUtils.createFormulaEvaluator(workbook);

        final Collection<?> rowObjects = parseSheet(workbook, sheetMeta, formulaEvaluator, validator);

        PxlWorkbookUtils.closeWorkbook(workbook);

        return rowObjects;
    }

    /**
     * Resolves the actual sheet index/name for every sheet meta and enforces the sheet-count limit. (import)
     *
     * @param workbook     the open POI workbook to inspect
     * @param workbookMeta the resolved import metadata for the workbook
     * @param sheetMetas   the sheet metas to resolve against the workbook's sheets
     * @throws PxlNullPointerException if {@code workbook} is {@code null}
     * @throws PxlDataException        if the workbook exceeds the maximum sheet count, or a required or duplicate sheet check fails
     */
    private static void resolveAllSheetsFromWorkbook(final Workbook workbook,
                                                     final PxlImportWorkbookMeta workbookMeta,
                                                     final List<PxlImportSheetMeta> sheetMetas)
            throws PxlNullPointerException, PxlDataException {

        final int numOfSheets = workbook.getNumberOfSheets();
        final int maxNumOfSheets = workbookMeta.getImportFileFormat().getMaxImportSheets();
        if (numOfSheets > maxNumOfSheets) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_SHEET_COUNT_EXCEEDED, String.valueOf(maxNumOfSheets)));
        }

        for (final PxlImportSheetMeta sheetMeta : sheetMetas) {
            resolveSheetFromWorkbook(workbook, sheetMeta);
        }
    }

    /**
     * Locates the physical sheet matching the sheet meta's name (or candidate names) and records its index. (import)
     * Sheet names are compared after whitespace removal, ignoring case.
     *
     * @param workbook  the open POI workbook to inspect
     * @param sheetMeta the sheet meta to resolve; its actual index/name are set on match
     * @throws PxlNullPointerException if {@code workbook} is {@code null}
     * @throws PxlDataException        if the sheet appears more than once, or a required sheet is missing
     */
    private static void resolveSheetFromWorkbook(final Workbook workbook,
                                                 final PxlImportSheetMeta sheetMeta)
            throws PxlNullPointerException, PxlDataException {

        if (!sheetMeta.isImportEnabled()) {
            return;
        }

        final String importSheetName = sheetMeta.getActualImportSheetName();
        final List<String> sheetNames = StringUtils.isBlank(importSheetName) ?
                sheetMeta.getCandidateSheetNames() :
                Collections.singletonList(importSheetName);

        final int numOfSheets = workbook.getNumberOfSheets();
        for (int workbookSheetIndex = 0; workbookSheetIndex < numOfSheets; workbookSheetIndex++) {
            // Only the name is compared here, so it is read without reaching for the sheet: on a streaming
            // workbook getSheetAt opens the sheet's own part to build a reader for it, which would mean opening
            // every sheet in the workbook to resolve one. A missing name is blank and falls through below.
            final String workbookSheetName = StringUtils.deleteWhitespace(workbook.getSheetName(workbookSheetIndex));
            if (StringUtils.isBlank(workbookSheetName)) {
                continue;
            }

            if (matchesSheetName(sheetNames, workbookSheetName)) {
                if (sheetMeta.getActualImportSheetIndex() >= 0) {
                    throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_SHEET_DUPLICATE, sheetNames));
                }

                sheetMeta.setActualImportSheetIndex(workbookSheetIndex);
                sheetMeta.setActualImportSheetName(workbookSheetName);
                // break;
            }
        }

        if ((sheetMeta.isRequired()) && (sheetMeta.getActualImportSheetIndex() < 0)) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_SHEET_NOT_FOUND, sheetNames, PxlWorkbookUtils.getSheetNamesOfWorkbook(workbook)));
        }
    }

    /**
     * Reads one resolved sheet row by row and binds each data row into a row object. (import)
     * Computes the 0-based header/data row bounds from the 1-based meta values, reads the header row to
     * resolve column indexes, populates the {@code @PxlRowIndex} field, skips empty and hidden rows,
     * optionally validates each row, and finally checks column uniqueness.
     *
     * @param workbook         the open POI workbook to read from
     * @param sheetMeta        the resolved sheet meta
     * @param formulaEvaluator the formula evaluator used to resolve formula cells (may be {@code null})
     * @param validator        optional bean validator applied when data validation is enabled (may be {@code null})
     * @return the collection of populated row objects, or {@code null} when the sheet is disabled, has no columns, or was not found
     * @throws PxlNullPointerException if {@code workbook} or {@code sheetMeta} is {@code null}
     * @throws PxlReflectionException  if instantiating a class or reading/writing a field fails
     * @throws PxlArgumentException    if the {@code @PxlRowIndex} field type is unsupported
     * @throws PxlDataException        if the header row is missing or a limit is exceeded
     * @throws PxlValidationException  if a bean-validation constraint on a row object is violated
     * @throws PxlCellCodecException   if a cell value cannot be decoded
     */
    private static Collection<?> parseSheet(final Workbook workbook,
                                            final PxlImportSheetMeta sheetMeta,
                                            @Nullable final FormulaEvaluator formulaEvaluator,
                                            @Nullable final Validator validator)
            throws PxlNullPointerException, PxlReflectionException, PxlArgumentException, PxlDataException, PxlValidationException, PxlCellCodecException {

        PxlAssertSupport.notNull(workbook, "workbook");
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

        // Reference the sheet that uses the given name.
        final int importSheetIndex = sheetMeta.getActualImportSheetIndex();
        if (importSheetIndex < 0) {
            return null;
        }

        final String sheetName = sheetMeta.getActualImportSheetName();
        final Sheet sheet = workbook.getSheetAt(importSheetIndex);

        // NOTE: StreamingSheet's getFirstRowNum() is not supported.
        final int firstRowNum = (sheet instanceof StreamingSheet) ? 0 : sheet.getFirstRowNum();

        // NOTE: StreamingSheet's getLastRowNum() sometimes abnormally returns 0.
        final int lastRowNum = (sheet instanceof StreamingSheet)
                ? workbookMeta.getImportFileFormat().getMaxImportRows()
                : sheet.getLastRowNum() + 1;    // Add 1 to make it exclusive.

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

        /* sheet.getPhysicalNumberOfRows() */
        // In practice we import Excel files that already conform to the limit below, so the code below has no real effect.
        final int maxNumOfRows = workbookMeta.getImportFileFormat().getMaxImportRows();
        if (actualImportBoundDataRowIndex > maxNumOfRows) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_SHEET_ROW_COUNT_EXCEEDED, sheetName, String.valueOf(maxNumOfRows)));
        }

        sheetMeta.setActualImportHeaderRowIndex(actualImportHeaderRowIndex);
        sheetMeta.setActualImportOriginDataRowIndex(actualImportOriginDataRowIndex);
        sheetMeta.setActualImportBoundDataRowIndex(actualImportBoundDataRowIndex);

        // Obtain the row index field.
        final Field rowIndexField = getRowIndexField(rowClass);
        Class<?> rowIndexClass = null;
        if (Objects.nonNull(rowIndexField)) {
            rowIndexClass = rowIndexField.getType();
        }

        final boolean importDataValidation = workbookMeta.isImportDataValidation();
        final boolean importExcludeHiddenRows = sheetMeta.isImportExcludeHiddenRows();

        final Collection<Object> rowObjects = (Collection<Object>) PxlReflectionSupport.newClassInstance(rowCollectionClass);

        // Read each data row.
        boolean headerRowProcessed = false;
        for (final Row row : sheet) {

            final int rowIndex = row.getRowNum();

            if (rowIndex == actualImportHeaderRowIndex) {
                readHeaderRowFromSheet(sheet, row, sheetMeta);
                headerRowProcessed = true;

                final boolean noImportColumn = columnMetas.stream().allMatch(c -> c.getActualImportColumnIndex() < 0);
                if (noImportColumn) {
                    throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_SHEET_NO_HEADER_COLUMN, sheetName, String.valueOf(actualImportHeaderRowIndex + 1)));
                }
            } else if (rowIndex >= actualImportOriginDataRowIndex && rowIndex < actualImportBoundDataRowIndex) {

                if (importExcludeHiddenRows && row.getZeroHeight()) {
                    continue;
                }

                final Object rowObject = PxlReflectionSupport.newClassInstance(rowClass);

                if (Objects.nonNull(rowIndexField) && Objects.nonNull(rowIndexClass)) {
                    // Expose the 0-based POI row number as a 1-based spreadsheet row number.
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

                readDataRowFromSheet(sheet, row, sheetMeta, rowObject, formulaEvaluator);

                // Skip the row if it is ignorable.
                if (isIgnorableRow(columnMetas, rowObject)) {
                    continue;
                }

                if (importDataValidation && Objects.nonNull(validator)) {
                    validateBeanConstraints(validator, rowObject, sheetName, rowIndex);
                }

                rowObjects.add(rowObject);
            } else if (rowIndex >= actualImportBoundDataRowIndex) {
                break;
            }
        }

        // If the header row does not physically exist (empty sheet, misconfigured header row index, etc.), column index resolution
        // and required-column checks are never performed at all, so an empty result is silently returned. To match CSV behavior (which always processes a header), fail explicitly here.
        if (!headerRowProcessed) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_SHEET_NO_HEADER_ROW, sheetName, String.valueOf(actualImportHeaderRowIndex + 1)));
        }

        validateDataUniqueness(rowObjects, columnMetas, sheetName);

        return rowObjects;
    }

    /**
     * Reads the header row and resolves each column's actual index by matching its header cell text
     * against the column's candidate names. (import)
     * Header cell text is compared after whitespace removal; hidden columns are skipped when configured.
     *
     * @param sheet     the sheet being read
     * @param headerRow the physical header row
     * @param sheetMeta the resolved sheet meta whose column metas are updated with matched indexes/names
     * @throws PxlDataException if the header row has no cells, a column exceeds the column limit,
     *                          a column name is duplicated, or a required column is missing
     */
    private static void readHeaderRowFromSheet(final Sheet sheet,
                                               final Row headerRow,
                                               final PxlImportSheetMeta sheetMeta)
            throws PxlDataException {

        final PxlImportWorkbookMeta workbookMeta = sheetMeta.getWorkbookMeta();
        final List<PxlImportColumnMeta> columnMetas = sheetMeta.getImportColumnMetas();

        final String sheetName = sheet.getSheetName();
        final int headerRowIndex = headerRow.getRowNum();

        final short firstCellNum = headerRow.getFirstCellNum();
        final short lastCellNum = headerRow.getLastCellNum();
        if (firstCellNum < 0 || lastCellNum < 0) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_SHEET_NO_HEADER_ROW, sheetName, String.valueOf(headerRowIndex + 1)));
        }

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

        // In practice we import Excel files that already conform to the limit below, so the code below has no real effect.
        final int maxNumOfColumns = workbookMeta.getImportFileFormat().getMaxImportColumns();
        if (actualImportBoundDataColumnIndex > maxNumOfColumns) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_SHEET_COLUMN_COUNT_EXCEEDED, sheetName, String.valueOf(maxNumOfColumns)));
        }

        sheetMeta.setActualImportOriginDataColumnIndex(actualImportOriginDataColumnIndex);
        sheetMeta.setActualImportBoundDataColumnIndex(actualImportBoundDataColumnIndex);

        final boolean importExcludeHiddenColumns = sheetMeta.isImportExcludeHiddenColumns();

        for (final PxlImportColumnMeta columnMeta : columnMetas) {
            if (!columnMeta.isImportEnabled()) {
                continue;
            }

            final List<String> candidateColumnNames = columnMeta.getCandidateColumnNames();

            for (int importColumnIndex = actualImportOriginDataColumnIndex; importColumnIndex < actualImportBoundDataColumnIndex; importColumnIndex++) {

                if (importExcludeHiddenColumns && sheet.isColumnHidden(importColumnIndex)) {
                    continue;
                }

                final Cell cell = headerRow.getCell(importColumnIndex);
                if (Objects.isNull(cell)) {
                    continue;
                }

                final String columnName = StringUtils.deleteWhitespace(PxlCellUtils.getCellStringValue(cell, workbookMeta.getImportDataFormatterCache()));
                if (StringUtils.isBlank(columnName)) {
                    continue;
                }

                if (candidateColumnNames.contains(columnName)) {
                    if (columnMeta.getActualImportColumnIndex() >= 0) {
                        throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_COLUMN_DUPLICATE, sheetName, candidateColumnNames));
                    }

                    columnMeta.setActualImportColumnIndex(importColumnIndex);
                    columnMeta.setActualImportColumnName(columnName);
                    // break;
                }
            }

            if ((columnMeta.isRequired()) && (columnMeta.getActualImportColumnIndex() < 0)) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_COLUMN_NOT_FOUND, sheetName, candidateColumnNames));
            }
        }
    }

    /**
     * Reads one data row and decodes each mapped cell into the corresponding field of the row object. (import)
     * Blank cells are skipped; when merged-region handling is enabled, the cell is resolved through its merge.
     *
     * @param sheet            the sheet being read
     * @param row              the physical data row
     * @param sheetMeta        the resolved sheet meta providing the column metas
     * @param rowObject        the row object populated with decoded values
     * @param formulaEvaluator the formula evaluator used to resolve formula cells (may be {@code null})
     * @throws PxlCellCodecException if a cell value cannot be decoded into the target field type
     */
    private static void readDataRowFromSheet(final Sheet sheet,
                                             final Row row,
                                             final PxlImportSheetMeta sheetMeta,
                                             final Object rowObject,
                                             final FormulaEvaluator formulaEvaluator)
            throws PxlCellCodecException {

        final List<PxlImportColumnMeta> columnMetas = sheetMeta.getImportColumnMetas();

        final int rowIndex = row.getRowNum();
        final String sheetName = sheet.getSheetName();
        final boolean importEachCellOfMergedRegion = sheetMeta.isImportEachCellOfMergedRegion();

        for (final PxlImportColumnMeta columnMeta : columnMetas) {
            if (!columnMeta.isImportEnabled()) {
                continue;
            }

            final int importColumnIndex = columnMeta.getActualImportColumnIndex();
            if (importColumnIndex < 0) {
                continue;
            }

            final Cell cell = importEachCellOfMergedRegion
                    ? PxlCellUtils.getCellWithMerges(sheet, rowIndex, importColumnIndex)
                    : row.getCell(importColumnIndex);
            if (PxlCellUtils.isBlankCell(cell)) {
                continue;
            }

            try {
                final Object valueObject = PxlCellResolver.parseDataValueFromCell(cell, columnMeta, formulaEvaluator);
                final Field columnField = columnMeta.getColumnField();

                if (Objects.nonNull(valueObject)) {
                    PxlReflectionSupport.setFieldValue(columnField, rowObject, valueObject);
                }
            } catch (Exception e) {
                throw new PxlCellCodecException(sheetName, rowIndex, columnMeta.getActualImportColumnName(), importColumnIndex, e);
            }
        }
    }

}
