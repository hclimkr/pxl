package io.github.hclimkr.pxl.internal.core;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.exception.*;
import io.github.hclimkr.pxl.internal.codec.PxlCellResolver;
import io.github.hclimkr.pxl.internal.codec.PxlEnumCodec;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlExportColumnMeta;
import io.github.hclimkr.pxl.internal.meta.PxlExportSheetMeta;
import io.github.hclimkr.pxl.internal.meta.PxlExportWorkbookMeta;
import io.github.hclimkr.pxl.internal.support.*;
import io.github.hclimkr.pxl.option.PxlExportSheetOption;
import io.github.hclimkr.pxl.styler.PxlStyler;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import io.github.hclimkr.pxl.util.PxlColumnUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.validation.Validator;
import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Stream;

/**
 * Excel export routine
 */
public final class PxlCoreExcelExporter extends PxlAbstractExporter {

    /**
     * Prevents instantiation.
     */
    private PxlCoreExcelExporter() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Builds a workbook from an annotated workbook object. (export)
     * Optionally validates the object and its rows, resolves each {@code @PxlSheet} field into a sheet,
     * enforces the sheet-count limit, and triggers formula recalculation when any string-as-formula column exists.
     *
     * @param workbookObject the annotated workbook object whose sheet fields supply the row data
     * @param workbookMeta   the resolved export metadata for the workbook
     * @param validator      optional bean validator applied when data validation is enabled (may be {@code null})
     * @return the populated POI workbook
     * @throws PxlNullPointerException if {@code workbookObject} or {@code workbookMeta} is {@code null}
     * @throws PxlValidationException  if a bean-validation constraint on a row object is violated
     * @throws PxlDataException        if a limit is exceeded or there is no data to write
     * @throws PxlReflectionException  if instantiating a class or reading/writing a field fails
     * @throws PxlArgumentException    if a configuration value is invalid
     * @throws PxlCellCodecException   if a cell value cannot be encoded
     * @see <a href="https://poi.apache.org/components/spreadsheet/eval.html#sxssf">SXSSF formula evaluation</a>
     */
    public static Workbook buildWorkbook(final Object workbookObject,
                                         final PxlExportWorkbookMeta workbookMeta,
                                         @Nullable final Validator validator)
            throws PxlNullPointerException, PxlValidationException, PxlDataException, PxlReflectionException, PxlArgumentException, PxlCellCodecException {

        PxlAssertSupport.notNull(workbookObject, "workbookObject");
        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");

        final boolean exportDataValidation = workbookMeta.isExportDataValidation();
        if (exportDataValidation && Objects.nonNull(validator)) {
            validateDataConstraint(validator, workbookObject, null, null);
        }

        final List<PxlExportSheetOption> sheetOptions = workbookMeta.getExportSheetOptions();
        final Workbook workbook = workbookMeta.getWorkbook();

        final List<PxlExportSheetMeta> sheetMetas = PxlExportSheetMeta.makeExportSheetMetas(workbookObject.getClass(), workbookMeta, sheetOptions, false);
        workbookMeta.addExportSheetMetas(sheetMetas);

        final long numOfSheets = sheetMetas.stream().filter(PxlExportSheetMeta::isExportEnabled).count();
        final int maxNumOfSheets = workbookMeta.getExportFileFormat().getMaxExportSheets();
        if (numOfSheets > maxNumOfSheets) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_SHEET_COUNT_EXCEEDED, String.valueOf(maxNumOfSheets)));
        }

        for (final PxlExportSheetMeta sheetMeta : sheetMetas) {
            if (!sheetMeta.isExportEnabled()) {
                continue;
            }

            // Get the row data to write into the sheet.
            final Collection<?> rowObjects = getRowObjects(sheetMeta.getSheetField(), workbookObject);

            if (exportDataValidation && Objects.nonNull(validator) && Objects.nonNull(rowObjects)) {
                validateDataConstraint(validator, rowObjects, sheetMeta.getActualExportSheetName(), null);
            }

            buildSheet(sheetMeta, rowObjects);
        }

        if (workbook.getNumberOfSheets() < 1) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_NO_DATA));
        }

        if (workbookMeta.hasAnyExportStringAsFormulaColumn()) {
            if (workbook instanceof SXSSFWorkbook) {
                // Delegate re-calculation to Excel. The application will perform a full recalculation when the workbook is opened.
                workbook.setForceFormulaRecalculation(true);
            } else {
                final FormulaEvaluator formulaEvaluator = workbookMeta.getFormulaEvaluator();
                if (Objects.nonNull(formulaEvaluator)) {
                    formulaEvaluator.evaluateAll();
                }
            }
        }

        return workbook;
    }

    /**
     * Builds a workbook containing a single sheet from a collection of row objects. (export)
     * A wildcard sheet option (if present) is applied, and formula recalculation is triggered
     * when any string-as-formula column exists.
     *
     * @param sheetName    the name of the sheet to create
     * @param rowObjects   the row objects written as data rows
     * @param rowClass     the row class describing the column bindings
     * @param workbookMeta the resolved export metadata for the workbook
     * @param validator    optional bean validator applied when data validation is enabled (may be {@code null})
     * @return the populated POI workbook
     * @throws PxlNullPointerException if {@code sheetName}, {@code rowObjects}, {@code rowClass}, or {@code workbookMeta} is {@code null}
     * @throws PxlArgumentException    if {@code sheetName} is blank
     * @throws PxlValidationException  if a bean-validation constraint on a row object is violated
     * @throws PxlDataException        if there is no data to write
     * @throws PxlCellCodecException   if a cell value cannot be encoded
     * @throws PxlReflectionException  if instantiating a class or reading/writing a field fails
     * @see <a href="https://poi.apache.org/components/spreadsheet/eval.html#sxssf">SXSSF formula evaluation</a>
     */
    public static Workbook buildWorkbook(final String sheetName,
                                         final Collection<?> rowObjects,
                                         final Class<?> rowClass,
                                         final PxlExportWorkbookMeta workbookMeta,
                                         @Nullable final Validator validator)
            throws PxlNullPointerException, PxlArgumentException, PxlValidationException, PxlDataException, PxlCellCodecException, PxlReflectionException {

        PxlAssertSupport.notBlank(sheetName, "sheetName");
        PxlAssertSupport.notNull(rowObjects, "rowObjects");
        PxlAssertSupport.notNull(rowClass, "rowClass");
        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");

        final PxlExportSheetOption sheetOption = Optional.ofNullable(workbookMeta.getExportSheetOptions())
                .flatMap(options -> options.stream()
                        .filter(o -> StringUtils.equals(o.getFieldName(), PxlConstants.SHEET_FIELD_NAME_WILD_CARD))
                        .findFirst())
                .orElse(null);

        final boolean exportDataValidation = workbookMeta.isExportDataValidation();
        if (exportDataValidation && Objects.nonNull(validator)) {
            validateDataConstraint(validator, rowObjects, null, null);
        }

        final Workbook workbook = workbookMeta.getWorkbook();

        final Class<?> rowCollectionClass = rowObjects.getClass();
        final PxlExportSheetMeta sheetMeta = PxlExportSheetMeta.makeExportSheetMeta(sheetName, rowCollectionClass, rowClass, workbookMeta, sheetOption);
        workbookMeta.addExportSheetMeta(sheetMeta);

        if (!sheetMeta.isExportEnabled()) {
            return workbook;
        }

        buildSheet(sheetMeta, rowObjects);

        if (workbook.getNumberOfSheets() < 1) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_NO_DATA));
        }

        if (workbookMeta.hasAnyExportStringAsFormulaColumn()) {
            if (workbook instanceof SXSSFWorkbook) {
                // Delegate re-calculation to Excel. The application will perform a full recalculation when the workbook is opened.
                workbook.setForceFormulaRecalculation(true);
            } else {
                final FormulaEvaluator formulaEvaluator = workbookMeta.getFormulaEvaluator();
                if (Objects.nonNull(formulaEvaluator)) {
                    formulaEvaluator.evaluateAll();
                }
            }
        }

        return workbook;
    }

    /**
     * Builds a workbook with one sheet per given name/collection pair. (export)
     * Validates that the name, data, and row-class lists are the same size and that sheet names are unique
     * (after safe-name normalization), enforces the sheet-count limit, and triggers formula recalculation
     * when any string-as-formula column exists.
     *
     * @param sheetNames   the sheet names, one per sheet
     * @param sheetObjects the per-sheet row-object collections, aligned with {@code sheetNames}
     * @param rowClasses   the per-sheet row classes, aligned with {@code sheetNames}
     * @param workbookMeta the resolved export metadata for the workbook
     * @param validator    optional bean validator applied when data validation is enabled (may be {@code null})
     * @return the populated POI workbook
     * @throws PxlNullPointerException if {@code sheetNames}, {@code sheetObjects}, {@code rowClasses}, or {@code workbookMeta} is {@code null}
     * @throws PxlDataException        if the list sizes differ, a sheet name is duplicated, an element is {@code null}, a limit is exceeded, or there is no data to write
     * @throws PxlValidationException  if a bean-validation constraint on a row object is violated
     * @throws PxlArgumentException    if a configuration value is invalid
     * @throws PxlCellCodecException   if a cell value cannot be encoded
     * @throws PxlReflectionException  if instantiating a class or reading/writing a field fails
     * @see <a href="https://poi.apache.org/components/spreadsheet/eval.html#sxssf">SXSSF formula evaluation</a>
     */
    public static Workbook buildWorkbook(final List<String> sheetNames,
                                         final List<Collection<?>> sheetObjects,
                                         final List<Class<?>> rowClasses,
                                         final PxlExportWorkbookMeta workbookMeta,
                                         @Nullable final Validator validator)
            throws PxlNullPointerException, PxlDataException, PxlValidationException, PxlArgumentException, PxlCellCodecException, PxlReflectionException {

        PxlAssertSupport.notNull(sheetNames, "sheetNames");
        PxlAssertSupport.notNull(sheetObjects, "sheetObjects");
        PxlAssertSupport.notNull(rowClasses, "rowClasses");
        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");

        if (PxlCollectionUtils.size(sheetNames) != PxlCollectionUtils.size(sheetObjects)) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_SHEET_NAME_OBJECT_COUNT_MISMATCH));
        }

        if (PxlCollectionUtils.size(sheetNames) != PxlCollectionUtils.size(rowClasses)) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_SHEET_NAME_ROW_CLASS_COUNT_MISMATCH));
        }

        final Set<String> duplicatedSheetNames = PxlWorkbookSupport.findDuplicateSheetNames(sheetNames);
        if (!duplicatedSheetNames.isEmpty()) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_DUPLICATE_SHEET_NAME, duplicatedSheetNames));
        }

        for (int index = 0; index < PxlCollectionUtils.size(sheetObjects); index++) {
            if (Objects.isNull(PxlCollectionUtils.get(sheetObjects, index))) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_SHEET_DATA_NULL, PxlCollectionUtils.get(sheetNames, index)));
            }
            if (Objects.isNull(PxlCollectionUtils.get(rowClasses, index))) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_SHEET_ROW_CLASS_NULL, PxlCollectionUtils.get(sheetNames, index)));
            }
        }

        final boolean exportDataValidation = workbookMeta.isExportDataValidation();
        if (exportDataValidation && Objects.nonNull(validator)) {
            for (final Collection<?> rowObjects : sheetObjects) {
                validateDataConstraint(validator, rowObjects, null, null);
            }
        }

        final Workbook workbook = workbookMeta.getWorkbook();

        final int numOfSheets = PxlCollectionUtils.size(sheetObjects);
        final int maxNumOfSheets = workbookMeta.getExportFileFormat().getMaxExportSheets();
        if (numOfSheets > maxNumOfSheets) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_SHEET_COUNT_EXCEEDED, String.valueOf(maxNumOfSheets)));
        }

        for (int index = 0; index < PxlCollectionUtils.size(sheetNames); index++) {
            final String sheetName = PxlCollectionUtils.get(sheetNames, index);
            final Collection<?> rowObjects = PxlCollectionUtils.get(sheetObjects, index);
            final Class<?> rowClass = PxlCollectionUtils.get(rowClasses, index);
            final PxlExportSheetOption sheetOption = workbookMeta.getExportSheetOption(index);

            final Class<?> rowCollectionClass = rowObjects.getClass();
            final PxlExportSheetMeta sheetMeta = PxlExportSheetMeta.makeExportSheetMeta(sheetName, rowCollectionClass, rowClass, workbookMeta, sheetOption);
            workbookMeta.addExportSheetMeta(sheetMeta);

            if (!sheetMeta.isExportEnabled()) {
                continue;
            }

            buildSheet(sheetMeta, rowObjects);
        }

        if (workbook.getNumberOfSheets() < 1) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_NO_DATA));
        }

        if (workbookMeta.hasAnyExportStringAsFormulaColumn()) {
            if (workbook instanceof SXSSFWorkbook) {
                // Delegate re-calculation to Excel. The application will perform a full recalculation when the workbook is opened.
                workbook.setForceFormulaRecalculation(true);
            } else {
                final FormulaEvaluator formulaEvaluator = workbookMeta.getFormulaEvaluator();
                if (Objects.nonNull(formulaEvaluator)) {
                    formulaEvaluator.evaluateAll();
                }
            }
        }

        return workbook;
    }

    /**
     * Builds a sample (template) workbook from an annotated workbook class: each sheet has a header row and a
     * single sample data row filled from each column's {@code exportSample} value. (export)
     * Each enabled {@code @PxlSheet} field becomes a sheet containing a header row and one sample data row.
     *
     * @param workbookClass the annotated workbook class describing the sheets and columns
     * @param workbookMeta  the resolved export metadata for the workbook
     * @return the populated sample POI workbook
     * @throws PxlNullPointerException if {@code workbookClass} or {@code workbookMeta} is {@code null}
     * @throws PxlDataException        if the sheet-count limit is exceeded or there is no sheet to create
     * @throws PxlReflectionException  if instantiating a class or reading/writing a field fails
     * @throws PxlArgumentException    if a configuration value is invalid
     * @throws PxlCellCodecException   if a cell value cannot be encoded
     */
    public static Workbook buildSampleWorkbook(final Class<?> workbookClass,
                                               final PxlExportWorkbookMeta workbookMeta)
            throws PxlNullPointerException, PxlDataException, PxlReflectionException, PxlArgumentException, PxlCellCodecException {

        PxlAssertSupport.notNull(workbookClass, "workbookClass");
        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");

        final List<PxlExportSheetOption> sheetOptions = workbookMeta.getExportSheetOptions();
        final Workbook workbook = workbookMeta.getWorkbook();

        final List<PxlExportSheetMeta> sheetMetas = PxlExportSheetMeta.makeExportSheetMetas(workbookClass, workbookMeta, sheetOptions, true);
        workbookMeta.addExportSheetMetas(sheetMetas);

        final long numOfSheets = sheetMetas.stream().filter(PxlExportSheetMeta::isExportSampleEnabled).count();
        final int maxNumOfSheets = workbookMeta.getExportFileFormat().getMaxExportSheets();
        if (numOfSheets > maxNumOfSheets) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_SHEET_COUNT_EXCEEDED, String.valueOf(maxNumOfSheets)));
        }

        for (final PxlExportSheetMeta sheetMeta : sheetMetas) {
            if (!sheetMeta.isExportSampleEnabled()) {
                continue;
            }

            buildSampleSheet(sheetMeta);
        }

        if (workbook.getNumberOfSheets() < 1) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_NO_DATA));
        }

        return workbook;
    }

    /**
     * Builds a sample (template) workbook with one sheet per given name/class pair; each sheet has a header row
     * and a single sample data row filled from each column's {@code exportSample} value. (export)
     * Validates that the name and row-class lists are the same size and that sheet names are unique
     * (after safe-name normalization); each sheet gets a header row and one sample data row.
     *
     * @param sheetNames   the sheet names, one per sheet
     * @param rowClasses   the per-sheet row classes, aligned with {@code sheetNames}
     * @param workbookMeta the resolved export metadata for the workbook
     * @return the populated sample POI workbook
     * @throws PxlNullPointerException if {@code sheetNames}, {@code rowClasses}, or {@code workbookMeta} is {@code null}
     * @throws PxlDataException        if the list sizes differ, a sheet name is duplicated, a row class is {@code null}, the sheet-count limit is exceeded, or there is no sheet to create
     * @throws PxlArgumentException    if a configuration value is invalid
     * @throws PxlReflectionException  if instantiating a class or reading/writing a field fails
     * @throws PxlCellCodecException   if a cell value cannot be encoded
     */
    public static Workbook buildSampleWorkbook(final List<String> sheetNames,
                                               final List<Class<?>> rowClasses,
                                               final PxlExportWorkbookMeta workbookMeta)
            throws PxlNullPointerException, PxlDataException, PxlArgumentException, PxlReflectionException, PxlCellCodecException {

        PxlAssertSupport.notNull(sheetNames, "sheetNames");
        PxlAssertSupport.notNull(rowClasses, "rowClasses");
        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");

        if (PxlCollectionUtils.size(sheetNames) != PxlCollectionUtils.size(rowClasses)) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_SHEET_NAME_ROW_CLASS_COUNT_MISMATCH));
        }

        final Set<String> duplicatedSheetNames = PxlWorkbookSupport.findDuplicateSheetNames(sheetNames);
        if (!duplicatedSheetNames.isEmpty()) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_DUPLICATE_SHEET_NAME, duplicatedSheetNames));
        }

        for (int index = 0; index < PxlCollectionUtils.size(rowClasses); index++) {
            if (Objects.isNull(PxlCollectionUtils.get(rowClasses, index))) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_SHEET_ROW_CLASS_NULL, PxlCollectionUtils.get(sheetNames, index)));
            }
        }

        final Workbook workbook = workbookMeta.getWorkbook();

        final int numOfSheets = PxlCollectionUtils.size(sheetNames);
        final int maxNumOfSheets = workbookMeta.getExportFileFormat().getMaxExportSheets();
        if (numOfSheets > maxNumOfSheets) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_SHEET_COUNT_EXCEEDED, String.valueOf(maxNumOfSheets)));
        }

        for (int index = 0; index < PxlCollectionUtils.size(sheetNames); index++) {
            final String sheetName = PxlCollectionUtils.get(sheetNames, index);
            final Class<?> rowClass = PxlCollectionUtils.get(rowClasses, index);
            final PxlExportSheetOption sheetOption = workbookMeta.getExportSheetOption(index);

            final PxlExportSheetMeta sheetMeta = PxlExportSheetMeta.makeExportSheetMeta(sheetName, null, rowClass, workbookMeta, sheetOption);
            workbookMeta.addExportSheetMeta(sheetMeta);

            if (!sheetMeta.isExportSampleEnabled()) {
                continue;
            }

            buildSampleSheet(sheetMeta);
        }

        if (workbook.getNumberOfSheets() < 1) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_NO_DATA));
        }

        return workbook;
    }

    /**
     * Builds one logical sheet, splitting into multiple physical sheets when a grouping field is set. (export)
     * Computes the 0-based header/data row bounds from the 1-based meta values, resolves the column metas,
     * and enforces the row- and sheet-count limits (accounting for group expansion).
     *
     * @param sheetMeta  the resolved sheet meta
     * @param rowObjects the row objects to write (may be {@code null} or empty)
     * @throws PxlDataException        if a row- or sheet-count limit is exceeded
     * @throws PxlCellCodecException   if a cell value cannot be encoded
     * @throws PxlReflectionException  if reading a field or resolving a generic type fails
     * @throws PxlNullPointerException if a required argument is null
     * @throws PxlArgumentException    if a configuration value is invalid
     */
    private static void buildSheet(final PxlExportSheetMeta sheetMeta,
                                   final Collection<?> rowObjects)
            throws PxlDataException, PxlCellCodecException, PxlReflectionException, PxlNullPointerException, PxlArgumentException {

        final PxlExportWorkbookMeta workbookMeta = sheetMeta.getWorkbookMeta();

        final int defaultHeaderRowIndex = 0;
        final int numOfObjects = PxlCollectionUtils.size(rowObjects);

        final int numOfSheets = workbookMeta.getWorkbook().getNumberOfSheets();
        final int maxNumOfSheets = workbookMeta.getExportFileFormat().getMaxExportSheets();

        int actualExportHeaderRowIndex = sheetMeta.getExportHeaderRowIndex();
        if (actualExportHeaderRowIndex == PxlConstants.DEFAULT_EXPORT_HEADER_ROW_INDEX) {
            actualExportHeaderRowIndex = defaultHeaderRowIndex;
        } else {
            actualExportHeaderRowIndex -= 1;  // Specified as 1-based, so convert to 0-based.
            actualExportHeaderRowIndex = Math.max(actualExportHeaderRowIndex, defaultHeaderRowIndex);
        }

        int actualExportOriginDataRowIndex = sheetMeta.getExportFirstDataRowIndex();
        if (actualExportOriginDataRowIndex == PxlConstants.DEFAULT_EXPORT_FIRST_DATA_ROW_INDEX) {
            actualExportOriginDataRowIndex = actualExportHeaderRowIndex + 1;
        } else {
            actualExportOriginDataRowIndex -= 1;  // Specified as 1-based, so convert to 0-based.
            actualExportOriginDataRowIndex = Math.max(actualExportOriginDataRowIndex, actualExportHeaderRowIndex + 1);
        }

        int actualExportBoundDataRowIndex = sheetMeta.getExportLastDataRowIndex();
        if (actualExportBoundDataRowIndex == PxlConstants.DEFAULT_EXPORT_LAST_DATA_ROW_INDEX) {
            actualExportBoundDataRowIndex = actualExportOriginDataRowIndex + numOfObjects;
        } else {
            actualExportBoundDataRowIndex -= 1;  // Specified as 1-based, so convert to 0-based.
            actualExportBoundDataRowIndex += 1;  // Add 1 to use it as an exclusive bound.
            actualExportBoundDataRowIndex = Math.min(actualExportBoundDataRowIndex, actualExportOriginDataRowIndex + numOfObjects);
        }

        final int maxNumOfRows = workbookMeta.getExportFileFormat().getMaxExportRows();
        final Field groupingField = sheetMeta.getExportGroupingField();
        final boolean isGrouping = Objects.nonNull(groupingField) && PxlCollectionUtils.isNotEmpty(rowObjects);
        if (!isGrouping && actualExportBoundDataRowIndex > maxNumOfRows) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_SHEET_ROW_COUNT_EXCEEDED, sheetMeta.getActualExportSheetName(), String.valueOf(maxNumOfRows)));
        }

        sheetMeta.setActualExportHeaderRowIndex(actualExportHeaderRowIndex);
        sheetMeta.setActualExportOriginDataRowIndex(actualExportOriginDataRowIndex);
        sheetMeta.setActualExportBoundDataRowIndex(actualExportBoundDataRowIndex);

        // Get the column information.
        final List<PxlExportColumnMeta> columnMetas = PxlExportColumnMeta.makeExportColumnMetas(sheetMeta, false);
        sheetMeta.addExportColumnMetas(columnMetas);

        // Do nothing if there are no columns.
        if (PxlCollectionUtils.isEmpty(columnMetas)) {
            return;
        }

        if (isGrouping) {
            // Preserve order with a LinkedHashMap so that group sheets are created in data appearance (insertion) order.
            final Map<Object, List<Object>> groupMap = new LinkedHashMap<>();
            for (final Object object : rowObjects) {
                final Object key = Objects.isNull(object) ? null : PxlReflectionSupport.getFieldValue(groupingField, object);
                groupMap.computeIfAbsent(key, k -> new ArrayList<>()).add(object);
            }

            // A group export creates as many sheets as the groupMap size (there are no empty groups),
            // so check the limit against (current sheet count + group count) before creation. (The pre-check above does not account for group expansion.)
            if (numOfSheets + groupMap.size() > maxNumOfSheets) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_SHEET_COUNT_EXCEEDED, String.valueOf(maxNumOfSheets)));
            }

            for (Object key : groupMap.keySet()) {
                // A null group key (value is null or reflection failed) is labeled with a placeholder ("(ungrouped)").
                final Object keyLabel = Objects.isNull(key) ? "(ungrouped)" : key;
                final String desiredSheetName = sheetMeta.getActualExportSheetName() + " - " + keyLabel;
                // If different group keys collide on the same safe sheet name (invalid-char replacement / 31-char truncation), createSheet throws,
                // so make it unique with a suffix so it does not clash with the workbook's existing sheet names.
                final String uniqueSheetName = PxlWorkbookSupport.makeUniqueSafeSheetName(workbookMeta.getWorkbook(), desiredSheetName);

                final List<Object> groupRowObjects = groupMap.get(key);
                final int groupBoundDataRowIndex = Math.min(actualExportBoundDataRowIndex, actualExportOriginDataRowIndex + PxlCollectionUtils.size(groupRowObjects));
                if (groupBoundDataRowIndex > maxNumOfRows) {
                    throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_EXPORT_SHEET_ROW_COUNT_EXCEEDED, uniqueSheetName, String.valueOf(maxNumOfRows)));
                }

                buildSheetInternally(sheetMeta, uniqueSheetName, groupRowObjects);
            }
        } else {
            // Earlier group expansion may have already reached the workbook limit, so check before creating a non-group sheet too.
            if (numOfSheets + 1 > maxNumOfSheets) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_SHEET_COUNT_EXCEEDED, String.valueOf(maxNumOfSheets)));
            }

            buildSheetInternally(sheetMeta, sheetMeta.getActualExportSheetName(), rowObjects);
        }
    }

    /**
     * Creates a single physical sheet and writes its header and data rows. (export)
     * Honors {@code exportIfNull}/{@code exportIfEmpty}, and temporarily narrows the sheet's data bound to the
     * actually written rows so that post-processing (auto-filter, dropdowns, widths) matches the written range.
     *
     * @param sheetMeta  the resolved sheet meta providing the column metas and row bounds
     * @param sheetName  the safe name of the physical sheet to create
     * @param rowObjects the row objects to write into this physical sheet
     * @throws PxlCellCodecException if a cell value cannot be encoded
     */
    private static void buildSheetInternally(final PxlExportSheetMeta sheetMeta,
                                             final String sheetName,
                                             final Collection<?> rowObjects)
            throws PxlCellCodecException {

        final PxlExportWorkbookMeta workbookMeta = sheetMeta.getWorkbookMeta();
        final List<PxlExportColumnMeta> columnMetas = sheetMeta.getExportColumnMetas();

        final boolean exportIfNull = sheetMeta.isExportIfNull();
        if (!exportIfNull && Objects.isNull(rowObjects)) {
            return;
        }

        final boolean exportIfEmpty = sheetMeta.isExportIfEmpty();
        if (!exportIfEmpty && Objects.nonNull(rowObjects) && rowObjects.isEmpty()) {
            return;
        }

        final int actualExportHeaderRowIndex = sheetMeta.getActualExportHeaderRowIndex();
        final int actualExportOriginDataRowIndex = sheetMeta.getActualExportOriginDataRowIndex();
        final int actualExportBoundDataRowIndex = sheetMeta.getActualExportBoundDataRowIndex();

        // Create a sheet using the given name.
        final Sheet sheet = workbookMeta.getWorkbook().createSheet(WorkbookUtil.createSafeSheetName(sheetName));

        preBuildRows(sheet, sheetMeta);

        // Create the header row as the first row.
        buildHeaderRow(sheet, columnMetas, actualExportHeaderRowIndex);

        // The end row (exclusive) of the data actually written to this (group) sheet. For a group export it can differ from the bound based on the total row count.
        int actualWrittenBoundDataRowIndex = actualExportOriginDataRowIndex;
        if (PxlCollectionUtils.isNotEmpty(rowObjects) && actualExportBoundDataRowIndex > actualExportOriginDataRowIndex) {
            // Write each row's data into a row.
            int rowIndex = actualExportOriginDataRowIndex;
            for (final Object rowObject : rowObjects) {
                buildDataRow(sheet, columnMetas, rowObject, rowIndex);
                rowIndex++;
                if (rowIndex >= actualExportBoundDataRowIndex) {
                    break;
                }
            }

            actualWrittenBoundDataRowIndex = rowIndex;
        }

        // So that the auto-filter/dropdown range matches the rows actually written, apply this sheet's real bound only during post-processing and then restore it.
        // (In a group export, sheetInfo's bound is based on the total row count, so it must be restored for the next group's loop.)
        sheetMeta.setActualExportBoundDataRowIndex(actualWrittenBoundDataRowIndex);
        postBuildRows(sheet, sheetMeta);
        sheetMeta.setActualExportBoundDataRowIndex(actualExportBoundDataRowIndex);
    }

    /**
     * Creates one sample (template) sheet with a header row and a single sample data row. (export)
     * The sample cell values come from each column's {@code exportSample}.
     *
     * @param sheetMeta the resolved sheet meta
     * @throws PxlDataException        if a limit is exceeded or there is no data
     * @throws PxlReflectionException  if instantiating a class or reading/writing a field fails
     * @throws PxlNullPointerException if a required argument is null
     * @throws PxlArgumentException    if a configuration value is invalid
     * @throws PxlCellCodecException   if a cell value cannot be encoded
     */
    private static void buildSampleSheet(final PxlExportSheetMeta sheetMeta)
            throws PxlDataException, PxlReflectionException, PxlNullPointerException, PxlArgumentException, PxlCellCodecException {

        final PxlExportWorkbookMeta workbookMeta = sheetMeta.getWorkbookMeta();

        final int defaultHeaderRowIndex = 0;

        int actualExportHeaderRowIndex = sheetMeta.getExportHeaderRowIndex();
        if (actualExportHeaderRowIndex == PxlConstants.DEFAULT_EXPORT_HEADER_ROW_INDEX) {
            actualExportHeaderRowIndex = defaultHeaderRowIndex;
        } else {
            actualExportHeaderRowIndex -= 1;  // Specified as 1-based, so convert to 0-based.
            actualExportHeaderRowIndex = Math.max(actualExportHeaderRowIndex, defaultHeaderRowIndex);
        }

        int actualExportOriginDataRowIndex = sheetMeta.getExportFirstDataRowIndex();
        if (actualExportOriginDataRowIndex == PxlConstants.DEFAULT_EXPORT_FIRST_DATA_ROW_INDEX) {
            actualExportOriginDataRowIndex = actualExportHeaderRowIndex + 1;
        } else {
            actualExportOriginDataRowIndex -= 1;  // Specified as 1-based, so convert to 0-based.
            actualExportOriginDataRowIndex = Math.max(actualExportOriginDataRowIndex, actualExportHeaderRowIndex + 1);
        }

        int actualExportBoundDataRowIndex = actualExportOriginDataRowIndex + 1;    // sample 1 row

        sheetMeta.setActualExportHeaderRowIndex(actualExportHeaderRowIndex);
        sheetMeta.setActualExportOriginDataRowIndex(actualExportOriginDataRowIndex);
        sheetMeta.setActualExportBoundDataRowIndex(actualExportBoundDataRowIndex); // exclusive

        // Get the column information.
        final List<PxlExportColumnMeta> columnMetas = PxlExportColumnMeta.makeExportColumnMetas(sheetMeta, true);
        sheetMeta.addExportColumnMetas(columnMetas);

        // Do nothing if there are no columns.
        if (PxlCollectionUtils.isEmpty(columnMetas)) {
            return;
        }

        // Create a sheet using the given name.
        final Sheet sheet = workbookMeta.getWorkbook().createSheet(WorkbookUtil.createSafeSheetName(sheetMeta.getActualExportSheetName()));

        preBuildRows(sheet, sheetMeta);

        // Create the header row as the first row.
        buildHeaderRow(sheet, columnMetas, actualExportHeaderRowIndex);

        buildSampleRow(sheet, columnMetas, actualExportOriginDataRowIndex);

        postBuildRows(sheet, sheetMeta);
    }

    /**
     * Writes the header cells for all mapped columns and freezes the pane below the header row. (export)
     * Required and optional columns use their respective header-cell stylers.
     *
     * @param sheet          the sheet being built
     * @param columnMetas    the per-column export metadata supplying the column names and stylers
     * @param headerRowIndex the 0-based index of the header row to create
     */
    private static void buildHeaderRow(final Sheet sheet,
                                       final List<PxlExportColumnMeta> columnMetas,
                                       final int headerRowIndex) {

        final Row row = sheet.createRow(headerRowIndex);

        for (final PxlExportColumnMeta columnMeta : columnMetas) {
            final int exportColumnIndex = columnMeta.getActualExportColumnIndex();
            if (exportColumnIndex < 0) {
                continue;
            }

            final Cell cell = row.createCell(exportColumnIndex);

            final Class<? extends PxlStyler> exportHeaderCellStyler = columnMeta.isRequired() ?
                    columnMeta.getExportColumnRequiredHeaderCellStyler() :
                    columnMeta.getExportColumnOptionalHeaderCellStyler();
            final CellStyle cellStyle = columnMeta.getWorkbookMeta().getCellStyle(exportHeaderCellStyler);
            if (Objects.nonNull(cellStyle)) {
                cell.setCellStyle(cellStyle);
            }

            cell.setCellValue(columnMeta.getActualExportColumnName());
        }

        // Freeze the header row.
        sheet.createFreezePane(0, headerRowIndex + 1);
    }

    /**
     * Writes one data row by reading each mapped field of the row object and encoding it into a cell. (export)
     * Each cell receives the column's data-cell style; {@code null} values are intentionally not skipped
     * so that {@code exportNullString} can take effect.
     *
     * @param sheet       the sheet being built
     * @param columnMetas the per-column export metadata
     * @param rowObject   the row object whose field values are written
     * @param rowIndex    the 0-based index of the row to create
     * @throws PxlCellCodecException if a field value cannot be read or encoded
     */
    private static void buildDataRow(final Sheet sheet,
                                     final List<PxlExportColumnMeta> columnMetas,
                                     final Object rowObject,
                                     final int rowIndex)
            throws PxlCellCodecException {

        final Row row = sheet.createRow(rowIndex);
        final String sheetName = sheet.getSheetName();

        for (final PxlExportColumnMeta columnMeta : columnMetas) {
            final int exportColumnIndex = columnMeta.getActualExportColumnIndex();
            if (exportColumnIndex < 0) {
                continue;
            }

            final Field columnField = columnMeta.getColumnField();

            Object cellObject;
            try {
                cellObject = PxlReflectionSupport.getFieldValue(columnField, rowObject);
            } catch (Exception e) {
                // continue;
                throw new PxlCellCodecException(sheetName, rowIndex, columnMeta.getActualExportColumnName(), exportColumnIndex, e);
            }

            // Not checked here in order to support exportNullString.
//            if (Objects.isNull(cellObject)) {
//                continue;
//            }

            final Cell cell = row.createCell(exportColumnIndex);

            final Class<? extends PxlStyler> exportColumnDataCellStyler = columnMeta.getExportColumnDataCellStyler();
            final CellStyle cellStyle = columnMeta.getWorkbookMeta().getCellStyle(exportColumnDataCellStyler);
            if (Objects.nonNull(cellStyle)) {
                cell.setCellStyle(cellStyle);
            }

            try {
                PxlCellResolver.buildDataCell(cell, cellObject, columnMeta);
            } catch (Exception e) {
                throw new PxlCellCodecException(sheetName, rowIndex, columnMeta.getActualExportColumnName(), exportColumnIndex, e);
            }
        }
    }

    /**
     * Writes a single sample data row using each column's {@code exportSample} value. (export)
     * Each cell receives the column's data-cell style.
     *
     * @param sheet       the sheet being built
     * @param columnMetas the per-column export metadata supplying the sample values
     * @param rowIndex    the 0-based index of the sample row to create
     * @throws PxlCellCodecException if a sample value cannot be encoded
     */
    private static void buildSampleRow(final Sheet sheet,
                                       final List<PxlExportColumnMeta> columnMetas,
                                       final int rowIndex)
            throws PxlCellCodecException {

        final Row row = sheet.createRow(rowIndex);
        final String sheetName = sheet.getSheetName();

        for (final PxlExportColumnMeta columnMeta : columnMetas) {
            final int exportColumnIndex = columnMeta.getActualExportColumnIndex();
            if (exportColumnIndex < 0) {
                continue;
            }

            //final Field columnField = columnMeta.getColumnField();

            final String cellObject = columnMeta.getExportSample();
            // Not checked here in order to support exportNullString.
//            if (StringUtils.isBlank(cellObject)) {
//                continue;
//            }

            final Cell cell = row.createCell(exportColumnIndex);

            final Class<? extends PxlStyler> exportColumnDataCellStyler = columnMeta.getExportColumnDataCellStyler();
            final CellStyle cellStyle = columnMeta.getWorkbookMeta().getCellStyle(exportColumnDataCellStyler);
            if (Objects.nonNull(cellStyle)) {
                cell.setCellStyle(cellStyle);
            }

            try {
                PxlCellResolver.buildDataCell(cell, cellObject, columnMeta);
            } catch (Exception e) {
                throw new PxlCellCodecException(sheetName, rowIndex, columnMeta.getActualExportColumnName(), exportColumnIndex, e);
            }
        }
    }

    /**
     * Reads the sheet field of the workbook object and returns its value as a collection. (export)
     *
     * @param sheetField     the {@code @PxlSheet} field to read
     * @param workbookObject the workbook object holding the field
     * @return the field value cast to {@link Collection}, or {@code null} if it is {@code null} or not a collection
     * @throws PxlReflectionException if the field value cannot be read
     */
    private static Collection<?> getRowObjects(final Field sheetField,
                                               final Object workbookObject)
            throws PxlReflectionException {

        final Object rowObjects = PxlReflectionSupport.getFieldValue(sheetField, workbookObject);

        if (Objects.isNull(rowObjects) || !(rowObjects instanceof Collection<?>)) {
            return null;
        }

        return (Collection<?>) rowObjects;
    }

    /**
     * Prepares the sheet before rows are written. (export)
     * Applies the default row height and, for streaming ({@link SXSSFSheet}) sheets, registers auto-sized
     * columns for width tracking.
     *
     * @param sheet     the sheet being built
     * @param sheetMeta the resolved sheet meta providing the column metas and row height
     */
    private static void preBuildRows(final Sheet sheet,
                                     final PxlExportSheetMeta sheetMeta) {

        final List<PxlExportColumnMeta> columnMetas = sheetMeta.getExportColumnMetas();

        final float exportRowHeightInPoints = sheetMeta.getExportRowHeightInPoints();
        if (exportRowHeightInPoints > 0.F) {
            sheet.setDefaultRowHeightInPoints(exportRowHeightInPoints);
        }

        if (sheet instanceof SXSSFSheet) {
            for (final PxlExportColumnMeta columnMeta : columnMetas) {
                final int exportColumnIndex = columnMeta.getActualExportColumnIndex();

                if (exportColumnIndex >= 0) {
                    final int exportColumnWidth = columnMeta.getExportColumnWidth();
                    if (exportColumnWidth == PxlConstants.EXPORT_AUTO_COLUMN_WIDTH) {
                        ((SXSSFSheet) sheet).trackColumnForAutoSizing(exportColumnIndex);
                    }
                }
            }
        }
    }

    /**
     * Finishes the sheet after all rows are written. (export)
     * Optionally sets the auto-filter, adds Enum/option dropdown lists, and adjusts column widths.
     *
     * @param sheet     the sheet being built
     * @param sheetMeta the resolved sheet meta providing the column metas and options
     */
    private static void postBuildRows(final Sheet sheet,
                                      final PxlExportSheetMeta sheetMeta) {

        final List<PxlExportColumnMeta> columnMetas = sheetMeta.getExportColumnMetas();

        // Set a filter on all columns.
        if (sheetMeta.isExportColumnFilter()) {
            showColumnFilter(sheet, sheetMeta);
        }

        // Set a dropdown list on Enum-typed columns.
        showDropDownList(sheet, sheetMeta);

        // Adjust the column widths.
        fitColumnWidth(sheet, columnMetas);
    }

    /**
     * Applies an auto-filter spanning the header row and the written data range. (export)
     *
     * @param sheet     the sheet being built
     * @param sheetMeta the resolved sheet meta providing the header/data row and column bounds
     */
    private static void showColumnFilter(final Sheet sheet,
                                         final PxlExportSheetMeta sheetMeta) {

        final int firstRowIndex = Math.min(sheetMeta.getActualExportHeaderRowIndex(), sheetMeta.getActualExportOriginDataRowIndex());
        final int lastRowIndex = sheetMeta.getActualExportBoundDataRowIndex() - 1;        // inclusive
        final int firstColumnIndex = sheetMeta.getActualExportOriginDataColumnIndex();
        final int lastColumnIndex = sheetMeta.getActualExportBoundDataColumnIndex() - 1;  // inclusive

        sheet.setAutoFilter(new CellRangeAddress(firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex));
    }

    /**
     * Adds data-validation dropdown lists over the written data range. (export)
     * For Enum columns configured with {@code SET}/{@code SORTED_SET}, the items come from the explicit
     * option items or, when absent, from the enum constants converted to strings (sorted for {@code SORTED_SET});
     * non-enum columns use their explicit option items when present.
     *
     * @param sheet     the sheet being built
     * @param sheetMeta the resolved sheet meta providing the column metas and data-row bounds
     */
    private static void showDropDownList(final Sheet sheet,
                                         final PxlExportSheetMeta sheetMeta) {

        final List<PxlExportColumnMeta> columnMetas = sheetMeta.getExportColumnMetas();

        final int firstRowIndex = sheetMeta.getActualExportOriginDataRowIndex();     // inclusive
        final int lastRowIndex = sheetMeta.getActualExportBoundDataRowIndex() - 1;    // inclusive

        if (firstRowIndex > lastRowIndex) {    // empty data row
            return;
        }

        for (final PxlExportColumnMeta columnMeta : columnMetas) {
            final int exportColumnIndex = columnMeta.getActualExportColumnIndex();
            if (exportColumnIndex < 0) {
                continue;
            }

            final String[] exportOptionItems = columnMeta.getExportOptionItems();
            final CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(firstRowIndex, lastRowIndex, exportColumnIndex, exportColumnIndex);

            final Field columnField = columnMeta.getColumnField();
            final Class<?> columnClass = columnField.getType();
            if (columnClass.isEnum()) {
                final PxlColumn.ExportEnumDropDownListStyle exportEnumDropDownListStyle = columnMeta.getExportEnumDropDownListStyle();

                if (PxlColumn.ExportEnumDropDownListStyle.SET.equals(exportEnumDropDownListStyle)
                        || PxlColumn.ExportEnumDropDownListStyle.SORTED_SET.equals(exportEnumDropDownListStyle)) {

                    if (ArrayUtils.isNotEmpty(exportOptionItems)) {
                        final String[] enumStrings = PxlColumn.ExportEnumDropDownListStyle.SORTED_SET.equals(exportEnumDropDownListStyle)
                                ? Arrays.stream(exportOptionItems).sorted().toArray(String[]::new)
                                : exportOptionItems;

                        PxlColumnSupport.setDropdownList(sheet, columnMeta, cellRangeAddressList, enumStrings);
                    } else {
                        final Object[] enumConstants = columnClass.getEnumConstants();
                        final Stream<String> enumStringsStream = Stream.of(enumConstants)
                                .map(e -> {
                                    try {
                                        return PxlEnumCodec.exportEnumToString(e, columnMeta.getExportCustomConverterMeta());
                                    } catch (PxlCellCodecException ignored) {
                                        return null;    // exclude items that fail to convert
                                    }
                                })
                                .filter(Objects::nonNull);
                        final String[] enumStrings = PxlColumn.ExportEnumDropDownListStyle.SORTED_SET.equals(exportEnumDropDownListStyle)
                                ? enumStringsStream.sorted(Comparator.naturalOrder()).toArray(String[]::new)
                                : enumStringsStream.toArray(String[]::new);

                        PxlColumnSupport.setDropdownList(sheet, columnMeta, cellRangeAddressList, enumStrings);
                    }
                }
            } else if (ArrayUtils.isNotEmpty(exportOptionItems)) {
                PxlColumnSupport.setDropdownList(sheet, columnMeta, cellRangeAddressList, exportOptionItems);
            }
        }
    }

    /**
     * Adjusts the width of each mapped column. (export)
     * String/collection picture columns get a fixed image-based width; other columns are auto-sized,
     * left untouched (negative width), or set to the explicit width (capped at 255 characters).
     *
     * @param sheet       the sheet being built
     * @param columnMetas the per-column export metadata providing widths and picture flags
     */
    private static void fitColumnWidth(final Sheet sheet,
                                       final List<PxlExportColumnMeta> columnMetas) {

        for (final PxlExportColumnMeta columnMeta : columnMetas) {
            final int exportColumnIndex = columnMeta.getActualExportColumnIndex();
            if (exportColumnIndex < 0) {
                continue;
            }

            final boolean isExportStringAsPicture = columnMeta.isExportStringAsPicture();
            if (isExportStringAsPicture) {
                final Field columnField = columnMeta.getColumnField();
                final Class<?> columnClass = columnField.getType();

                if (columnClass == String.class) {
                    final int pictureWidthPx = PxlConstants.EXPORT_PICTURE_SCREEN_WIDTH_IN_PIXELS;
                    final int picturePaddingPx = PxlConstants.EXPORT_PICTURE_SCREEN_PADDING_IN_PIXELS;
                    final int horizontalImageNum = 1;
                    final int columnWidthPx = (pictureWidthPx + picturePaddingPx) * horizontalImageNum + picturePaddingPx;
                    sheet.setColumnWidth(exportColumnIndex, (int) ((float) columnWidthPx / Units.DEFAULT_CHARACTER_WIDTH * 256.f));
                } else if (PxlClassSupport.isCollectionClass(columnClass)) {
                    final int pictureWidthPx = PxlConstants.EXPORT_PICTURE_SCREEN_WIDTH_IN_PIXELS;
                    final int picturePaddingPx = PxlConstants.EXPORT_PICTURE_SCREEN_PADDING_IN_PIXELS;
                    final int horizontalImageNum = PxlConstants.EXPORT_HORIZONTAL_NUMBER_OF_PICTURE;
                    final int columnWidthPx = (pictureWidthPx + picturePaddingPx) * horizontalImageNum + picturePaddingPx;
                    sheet.setColumnWidth(exportColumnIndex, (int) ((float) columnWidthPx / Units.DEFAULT_CHARACTER_WIDTH * 256.f));
                }
            } else {
                final int exportColumnWidth = columnMeta.getExportColumnWidth();
                if (exportColumnWidth == PxlConstants.EXPORT_AUTO_COLUMN_WIDTH) {
                    PxlColumnUtils.autoSizeColumns(sheet, exportColumnIndex);
                } else if (exportColumnWidth < 0) {
                    // don't set column width
                } else {
                    sheet.setColumnWidth(exportColumnIndex, Math.min(255 * 256, exportColumnWidth));
                }
            }
        }
    }

    /**
     * Applies the given horizontal/vertical alignment to every cell of one column. (export)
     *
     * @param sheet               the sheet being modified
     * @param columnIndex         the 0-based index of the column to align
     * @param horizontalAlignment the horizontal alignment to apply, or {@code null} to leave it unchanged
     * @param verticalAlignment   the vertical alignment to apply, or {@code null} to leave it unchanged
     * @deprecated unused; alignment is handled through cell stylers
     */
    @SuppressWarnings("unused")
    @Deprecated
    private static void setColumnRightAligned(final Sheet sheet,
                                              final int columnIndex,
                                              final HorizontalAlignment horizontalAlignment,
                                              final VerticalAlignment verticalAlignment) {

        final Workbook workbook = sheet.getWorkbook();
        final CellStyle cellStyle = workbook.createCellStyle();

        if (Objects.nonNull(horizontalAlignment)) {
            cellStyle.setAlignment(horizontalAlignment);
        }

        if (Objects.nonNull(verticalAlignment)) {
            cellStyle.setVerticalAlignment(verticalAlignment);
        }

        for (final Row row : sheet) {
            final Cell cell = row.getCell(columnIndex);
            if (Objects.nonNull(cell)) {
                cell.setCellStyle(cellStyle);
            }
        }
    }

}
