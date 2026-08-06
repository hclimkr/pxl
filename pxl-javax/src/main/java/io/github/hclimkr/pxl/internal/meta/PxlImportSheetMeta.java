package io.github.hclimkr.pxl.internal.meta;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlDataException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.exception.PxlReflectionException;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nContent;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.internal.support.PxlClassSupport;
import io.github.hclimkr.pxl.internal.support.PxlReflectionSupport;
import io.github.hclimkr.pxl.option.PxlImportColumnOption;
import io.github.hclimkr.pxl.option.PxlImportSheetOption;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import lombok.Getter;
import lombok.Setter;
import org.apache.commons.lang3.StringUtils;

import javax.validation.constraints.NotEmpty;
import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Excel sheet import metadata.
 */
@Getter
public final class PxlImportSheetMeta {

    private final PxlImportWorkbookMeta workbookMeta;

    private final Field sheetField;

    private final Class<? extends Collection<?>> rowCollectionClass;

    private final Class<?> rowClass;

    private List<String> candidateSheetNames;

    private final boolean importEnabled;

    private final boolean importOverrideSuperClassSheet;

    private final boolean importExcludeHiddenRows;

    private final boolean importExcludeHiddenColumns;

    private final boolean importEachCellOfMergedRegion;

    private final int importHeaderRowIndex;          // 1-based

    private final int importFirstDataRowIndex;      // 1-based inclusive

    private final int importLastDataRowIndex;       // 1-based inclusive

    private final int importFirstDataColumnIndex;   // 1-based inclusive

    private final int importLastDataColumnIndex;    // 1-based inclusive

    private final String importCsvCharset;          // CSV only; resolved against the workbook charset

    private final char importCsvDelimiter;          // CSV only; resolved against the workbook delimiter

    private final List<PxlImportColumnOption> importColumnOptions;

    private final List<PxlImportColumnMeta> importColumnMetas;

    private final boolean isRequired;

    @Setter
    private String actualImportSheetName;

    @Setter
    private int actualImportSheetIndex = -1;

    @Setter
    private int actualImportHeaderRowIndex = -1;         // 0-based

    @Setter
    private int actualImportOriginDataRowIndex = -1;    // 0-based inclusive

    @Setter
    private int actualImportBoundDataRowIndex = -1;     // 0-based exclusive

    @Setter
    private int actualImportOriginDataColumnIndex = -1; // 0-based inclusive

    @Setter
    private int actualImportBoundDataColumnIndex = -1;  // 0-based exclusive

    /**
     * Creates the resolved import metadata for one sheet, storing the merged option/annotation values.
     *
     * @param workbookMeta                  the enclosing workbook metadata
     * @param sheetField                    the {@link PxlSheet} Collection field to bind, or {@code null} for an ad-hoc sheet
     * @param rowCollectionClass            the concrete Collection type holding the rows
     * @param rowClass                      the row (element) class whose columns are bound
     * @param candidateSheetNames           the ordered candidate sheet names to match against
     * @param importEnabled                 whether this sheet is imported
     * @param importOverrideSuperClassSheet whether this sheet overrides a same-named super-class sheet
     * @param importExcludeHiddenRows       whether hidden rows are skipped
     * @param importExcludeHiddenColumns    whether hidden columns are skipped
     * @param importEachCellOfMergedRegion  whether each cell of a merged region is read individually
     * @param importHeaderRowIndex          the 1-based header row index
     * @param importFirstDataRowIndex       the 1-based inclusive first data row index
     * @param importLastDataRowIndex        the 1-based inclusive last data row index
     * @param importFirstDataColumnIndex    the 1-based inclusive first data column index
     * @param importLastDataColumnIndex     the 1-based inclusive last data column index
     * @param importCsvCharset              the resolved CSV charset for this sheet (CSV sources only)
     * @param importCsvDelimiter            the resolved CSV field delimiter for this sheet (CSV sources only)
     * @param importColumnOptions           the per-column import overrides
     * @param importColumnMetas             the (initially empty) column metadata list
     */
    private PxlImportSheetMeta(final PxlImportWorkbookMeta workbookMeta,
                               final Field sheetField,
                               final Class<? extends Collection<?>> rowCollectionClass,
                               final Class<?> rowClass,
                               final List<String> candidateSheetNames,
                               final boolean importEnabled,
                               final boolean importOverrideSuperClassSheet,
                               final boolean importExcludeHiddenRows,
                               final boolean importExcludeHiddenColumns,
                               final boolean importEachCellOfMergedRegion,
                               final int importHeaderRowIndex,
                               final int importFirstDataRowIndex,
                               final int importLastDataRowIndex,
                               final int importFirstDataColumnIndex,
                               final int importLastDataColumnIndex,
                               final String importCsvCharset,
                               final char importCsvDelimiter,
                               final List<PxlImportColumnOption> importColumnOptions,
                               final List<PxlImportColumnMeta> importColumnMetas) {

        this.workbookMeta = workbookMeta;

        this.sheetField = sheetField;
        this.rowCollectionClass = rowCollectionClass;
        this.rowClass = rowClass;
        this.candidateSheetNames = candidateSheetNames;
        this.importEnabled = importEnabled;
        this.importOverrideSuperClassSheet = importOverrideSuperClassSheet;
        this.importExcludeHiddenRows = importExcludeHiddenRows;
        this.importExcludeHiddenColumns = importExcludeHiddenColumns;
        this.importEachCellOfMergedRegion = importEachCellOfMergedRegion;
        this.importHeaderRowIndex = importHeaderRowIndex;
        this.importFirstDataRowIndex = importFirstDataRowIndex;
        this.importLastDataRowIndex = importLastDataRowIndex;
        this.importFirstDataColumnIndex = importFirstDataColumnIndex;
        this.importLastDataColumnIndex = importLastDataColumnIndex;
        this.importCsvCharset = importCsvCharset;
        this.importCsvDelimiter = importCsvDelimiter;
        this.importColumnOptions = importColumnOptions;
        this.importColumnMetas = importColumnMetas;

        this.isRequired = Objects.isNull(sheetField) || Objects.nonNull(sheetField.getAnnotation(NotEmpty.class));
    }

    /**
     * Returns a debug string of the candidate sheet names and, when present, the bound field name.
     *
     * @return the string representation
     */
    @Override
    public String toString() {

        if (Objects.nonNull(sheetField)) {
            return "[" + candidateSheetNames + ", " + sheetField.getName() + "]";
        } else {
            return "[" + candidateSheetNames + "]";
        }
    }

    /**
     * On import, collects the sheet metadata for the Excel file from the sheet options and the workbook class.
     * The sheet options take precedence over the workbook class.
     * Each {@link PxlSheet}-annotated Collection field becomes one sheet; sheet options (matched by field name,
     * falling back to the wildcard) override the annotation and i18n-resolved candidate names are computed.
     *
     * @param workbookClass the annotated workbook class whose fields are scanned for sheets; may be {@code null}
     * @param workbookMeta  the enclosing workbook metadata, supplying the i18n bundle
     * @param sheetOptions  per-field runtime sheet overrides; may be {@code null}
     * @return the collected sheet metadata list
     * @throws PxlNullPointerException if {@code workbookMeta} is {@code null}
     * @throws PxlReflectionException  if a sheet field's parameterized type cannot be resolved
     * @throws PxlDataException        if a row/column index is negative or inconsistent
     */
    public static List<PxlImportSheetMeta> makeImportSheetMetas(@Nullable final Class<?> workbookClass,
                                                                final PxlImportWorkbookMeta workbookMeta,
                                                                @Nullable final List<PxlImportSheetOption> sheetOptions)
            throws PxlNullPointerException, PxlReflectionException, PxlDataException {

        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");

//        final Field[] sheetFields = workbookClass.getDeclaredFields();
        final List<Field> sheetFields = PxlReflectionSupport.getAllFields(workbookClass);
        final List<PxlImportSheetMeta> sheetMetas = new ArrayList<>(PxlCollectionUtils.size(sheetFields));
        // Names differing only in case denote the same sheet on import, so an override must be recognized the same
        // way - a sheet declared as "EMPLOYEES" overrides a super-class sheet declared as "Employees", or else both
        // would bind the one physical sheet that matches either name.
        final Set<String> overriddenSheetNames = new TreeSet<>(String.CASE_INSENSITIVE_ORDER);

        for (final Field sheetField : sheetFields) {
            // Get the @PxlSheet annotation for the sheet field.
            final PxlSheet sheetAnnotation = sheetField.getAnnotation(PxlSheet.class);
            if (Objects.isNull(sheetAnnotation)) {
                continue;
            }

            // Ignore if it is not a Collection type.
            final Class<?> rowCollectionClass = sheetField.getType();
            if (!PxlClassSupport.isCollectionClass(rowCollectionClass)) {
                continue;
            }

            final Class<? extends Collection<?>> rowConcreteCollectionClass = PxlClassSupport.getConcreteCollectionClass(rowCollectionClass);

            // Get the generic class of the Collection.
            final Class<?> rowClass = PxlReflectionSupport.getParameterizedArgument0(sheetField);

            final PxlImportSheetOption sheetOption = Optional.ofNullable(sheetOptions)
                    .flatMap(options -> options.stream()
                            .filter(o -> StringUtils.equals(o.getFieldName(), sheetField.getName()))
                            .findFirst())
                    .orElseGet(() -> Optional.ofNullable(sheetOptions)
                            .flatMap(options -> options.stream()
                                    .filter(o -> StringUtils.equals(o.getFieldName(), PxlConstants.SHEET_FIELD_NAME_WILD_CARD))
                                    .findFirst())
                            .orElse(null));

            final List<String> candidateSheetNames = makeCandidateSheetNames(workbookMeta, sheetOption, sheetAnnotation, sheetField);

            // Ignore if the sheet name is not set.
            if (PxlCollectionUtils.isEmpty(candidateSheetNames)) {
                continue;
            }

            // Ignore if the sheet name is already overridden and in use.
            final boolean overriddenSheet = candidateSheetNames.stream()
                    .anyMatch(overriddenSheetNames::contains);
            if (overriddenSheet) {
                continue;
            }

            final boolean importEnabled = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getImportEnabled()))
                    .orElseGet(sheetAnnotation::importEnabled);

            final boolean importOverrideSuperClassSheet = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getImportOverrideSuperClassSheet()))
                    .orElseGet(sheetAnnotation::importOverrideSuperClassSheet);

            final boolean importExcludeHiddenRows = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getImportExcludeHiddenRows()))
                    .orElseGet(sheetAnnotation::importExcludeHiddenRows);

            final boolean importExcludeHiddenColumns = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getImportExcludeHiddenColumns()))
                    .orElseGet(sheetAnnotation::importExcludeHiddenColumns);

            final boolean importEachCellOfMergedRegion = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getImportEachCellOfMergedRegion()))
                    .orElseGet(sheetAnnotation::importEachCellOfMergedRegion);

            final int importHeaderRowIndex = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getImportHeaderRowIndex()))
                    .orElseGet(sheetAnnotation::importHeaderRowIndex);
            if (importHeaderRowIndex < 0) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "importHeaderRowIndex"));
            }

            final int importFirstDataRowIndex = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getImportFirstDataRowIndex()))
                    .orElseGet(sheetAnnotation::importFirstDataRowIndex);
            if (importFirstDataRowIndex < 0) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "importFirstDataRowIndex"));
            }

            final int importLastDataRowIndex = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getImportLastDataRowIndex()))
                    .orElseGet(sheetAnnotation::importLastDataRowIndex);
            if (importLastDataRowIndex < 0) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "importLastDataRowIndex"));
            }

            final int importFirstDataColumnIndex = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getImportFirstDataColumnIndex()))
                    .orElseGet(sheetAnnotation::importFirstDataColumnIndex);
            if (importFirstDataColumnIndex < 0) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "importFirstDataColumnIndex"));
            }

            final int importLastDataColumnIndex = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getImportLastDataColumnIndex()))
                    .orElseGet(sheetAnnotation::importLastDataColumnIndex);
            if (importLastDataColumnIndex < 0) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "importLastDataColumnIndex"));
            }

            if (importHeaderRowIndex != PxlConstants.DEFAULT_IMPORT_HEADER_ROW_INDEX
                    && importFirstDataRowIndex != PxlConstants.DEFAULT_IMPORT_FIRST_DATA_ROW_INDEX
                    && importFirstDataRowIndex <= importHeaderRowIndex) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_FIRST_DATA_ROW_GT_HEADER, candidateSheetNames, "importFirstDataRowIndex", "importHeaderRowIndex"));
            }

            if (importFirstDataRowIndex != PxlConstants.DEFAULT_IMPORT_FIRST_DATA_ROW_INDEX
                    && importLastDataRowIndex != PxlConstants.DEFAULT_IMPORT_LAST_DATA_ROW_INDEX
                    && importLastDataRowIndex < importFirstDataRowIndex) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_LAST_DATA_GE_FIRST_DATA, candidateSheetNames, "importLastDataRowIndex", "importFirstDataRowIndex"));
            }

            if (importFirstDataColumnIndex != PxlConstants.DEFAULT_IMPORT_FIRST_DATA_COLUMN_INDEX
                    && importLastDataColumnIndex != PxlConstants.DEFAULT_IMPORT_LAST_DATA_COLUMN_INDEX
                    && importLastDataColumnIndex < importFirstDataColumnIndex) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_LAST_DATA_GE_FIRST_DATA, candidateSheetNames, "importLastDataColumnIndex", "importFirstDataColumnIndex"));
            }

            final String importCsvCharset = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getImportCsvCharset()))
                    .filter(StringUtils::isNotBlank)
                    .orElseGet(() -> StringUtils.isNotBlank(sheetAnnotation.importCsvCharset()) ?
                            sheetAnnotation.importCsvCharset() :
                            workbookMeta.getImportCsvCharset());

            final char importCsvDelimiter = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getImportCsvDelimiter()))
                    .filter(delimiter -> delimiter != PxlConstants.UNSPECIFIED_IMPORT_CSV_DELIMITER)
                    .orElseGet(() -> sheetAnnotation.importCsvDelimiter() != PxlConstants.UNSPECIFIED_IMPORT_CSV_DELIMITER ?
                            sheetAnnotation.importCsvDelimiter() :
                            workbookMeta.getImportCsvDelimiter());

            final List<PxlImportColumnOption> importColumnOptions = Optional.ofNullable(sheetOption)
                    .map(option -> option.getImportColumnOptions())
                    .orElseGet(ArrayList::new);

            final List<PxlImportColumnMeta> importColumnMetas = new ArrayList<>();

            sheetField.setAccessible(true);
            sheetMetas.add(
                    new PxlImportSheetMeta(
                            workbookMeta,                   // workbookMeta
                            sheetField,                     // sheetField
                            rowConcreteCollectionClass,     // rowCollectionClass
                            rowClass,                       // rowClass
                            candidateSheetNames,            // candidateSheetNames
                            importEnabled,                  // importEnabled
                            importOverrideSuperClassSheet,  // importOverrideSuperClassSheet
                            importExcludeHiddenRows,        // importExcludeHiddenRows
                            importExcludeHiddenColumns,     // importExcludeHiddenColumns
                            importEachCellOfMergedRegion,   // importEachCellOfMergedRegion
                            importHeaderRowIndex,            // importHeaderRowIndex
                            importFirstDataRowIndex,        // importFirstDataRowIndex
                            importLastDataRowIndex,         // importLastDataRowIndex
                            importFirstDataColumnIndex,     // importFirstDataColumnIndex
                            importLastDataColumnIndex,      // importLastDataColumnIndex
                            importCsvCharset,               // importCsvCharset
                            importCsvDelimiter,             // importCsvDelimiter
                            importColumnOptions,            // importColumnOptions
                            importColumnMetas               // importColumnMetas
                    ));

            if (importEnabled && importOverrideSuperClassSheet) {
                overriddenSheetNames.addAll(candidateSheetNames);
            }
        }

        return sheetMetas;
    }

    /**
     * On import, collects the sheet metadata for the Excel file from the method parameters.
     * Used when importing an ad-hoc sheet: the candidate names are whitespace-stripped and blank entries dropped,
     * and the sheet option (when present) overrides the built-in defaults.
     *
     * @param candidateSheetNames the ordered candidate sheet names to match against; must not be empty after cleanup
     * @param rowCollectionClass  the Collection type to hold the rows; must be a Collection type
     * @param rowClass            the row (element) class whose columns are bound
     * @param workbookMeta        the enclosing workbook metadata
     * @param sheetOption         runtime overrides for this sheet; may be {@code null}
     * @return the single sheet metadata
     * @throws PxlNullPointerException if {@code candidateSheetNames}, {@code rowCollectionClass}, {@code rowClass}, or {@code workbookMeta} is {@code null}
     * @throws PxlArgumentException    if {@code candidateSheetNames} is empty
     * @throws PxlDataException        if {@code rowCollectionClass} is not a Collection type, no usable sheet name remains, or a row/column index is negative or inconsistent
     * @throws PxlReflectionException  if the row (element) type cannot be resolved
     */
    public static PxlImportSheetMeta makeImportSheetMeta(final List<String> candidateSheetNames,
                                                         final Class<?> rowCollectionClass,
                                                         final Class<?> rowClass,
                                                         final PxlImportWorkbookMeta workbookMeta,
                                                         @Nullable final PxlImportSheetOption sheetOption)
            throws PxlNullPointerException, PxlArgumentException, PxlDataException, PxlReflectionException {

        PxlAssertSupport.notEmpty(candidateSheetNames, "candidateSheetNames");
        PxlAssertSupport.notNull(rowCollectionClass, "rowCollectionClass");
        PxlAssertSupport.notNull(rowClass, "rowClass");
        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");

        if (!PxlClassSupport.isCollectionClass(rowCollectionClass)) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_NOT_COLLECTION_TYPE, rowCollectionClass.getSimpleName()));
        }

        final Class<? extends Collection<?>> rowConcreteCollectionClass = PxlClassSupport.getConcreteCollectionClass(rowCollectionClass);

        final List<String> sheetNames = candidateSheetNames.stream()
                .map(StringUtils::deleteWhitespace)
                .filter(StringUtils::isNotBlank)
                .collect(Collectors.toList());

        if (PxlCollectionUtils.isEmpty(sheetNames)) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_NAME_INVALID, candidateSheetNames));
        }

        final boolean importEnabled = true;

        final boolean importOverrideSuperClassSheet = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getImportOverrideSuperClassSheet()))
                .orElse(PxlConstants.DEFAULT_IMPORT_OVERRIDE_SUPER_CLASS_SHEET);

        final boolean importExcludeHiddenRows = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getImportExcludeHiddenRows()))
                .orElse(PxlConstants.DEFAULT_IMPORT_EXCLUDE_HIDDEN_ROWS);

        final boolean importExcludeHiddenColumns = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getImportExcludeHiddenColumns()))
                .orElse(PxlConstants.DEFAULT_IMPORT_EXCLUDE_HIDDEN_COLUMNS);

        final boolean importEachCellOfMergedRegion = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getImportEachCellOfMergedRegion()))
                .orElse(PxlConstants.DEFAULT_IMPORT_EACH_CELL_OF_MERGED_REGION);

        final int importHeaderRowIndex = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getImportHeaderRowIndex()))
                .orElse(PxlConstants.DEFAULT_IMPORT_HEADER_ROW_INDEX);
        if (importHeaderRowIndex < 0) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "importHeaderRowIndex"));
        }

        final int importFirstDataRowIndex = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getImportFirstDataRowIndex()))
                .orElse(PxlConstants.DEFAULT_IMPORT_FIRST_DATA_ROW_INDEX);
        if (importFirstDataRowIndex < 0) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "importFirstDataRowIndex"));
        }

        final int importLastDataRowIndex = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getImportLastDataRowIndex()))
                .orElse(PxlConstants.DEFAULT_IMPORT_LAST_DATA_ROW_INDEX);
        if (importLastDataRowIndex < 0) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "importLastDataRowIndex"));
        }

        final int importFirstDataColumnIndex = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getImportFirstDataColumnIndex()))
                .orElse(PxlConstants.DEFAULT_IMPORT_FIRST_DATA_COLUMN_INDEX);
        if (importFirstDataColumnIndex < 0) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "importFirstDataColumnIndex"));
        }

        final int importLastDataColumnIndex = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getImportLastDataColumnIndex()))
                .orElse(PxlConstants.DEFAULT_IMPORT_LAST_DATA_COLUMN_INDEX);
        if (importLastDataColumnIndex < 0) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "importLastDataColumnIndex"));
        }

        if (importHeaderRowIndex != PxlConstants.DEFAULT_IMPORT_HEADER_ROW_INDEX
                && importFirstDataRowIndex != PxlConstants.DEFAULT_IMPORT_FIRST_DATA_ROW_INDEX
                && importFirstDataRowIndex <= importHeaderRowIndex) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_FIRST_DATA_ROW_GT_HEADER, candidateSheetNames, "importFirstDataRowIndex", "importHeaderRowIndex"));
        }

        if (importFirstDataRowIndex != PxlConstants.DEFAULT_IMPORT_FIRST_DATA_ROW_INDEX
                && importLastDataRowIndex != PxlConstants.DEFAULT_IMPORT_LAST_DATA_ROW_INDEX
                && importLastDataRowIndex < importFirstDataRowIndex) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_LAST_DATA_GE_FIRST_DATA, candidateSheetNames, "importLastDataRowIndex", "importFirstDataRowIndex"));
        }

        if (importFirstDataColumnIndex != PxlConstants.DEFAULT_IMPORT_FIRST_DATA_COLUMN_INDEX
                && importLastDataColumnIndex != PxlConstants.DEFAULT_IMPORT_LAST_DATA_COLUMN_INDEX
                && importLastDataColumnIndex < importFirstDataColumnIndex) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_LAST_DATA_GE_FIRST_DATA, candidateSheetNames, "importLastDataColumnIndex", "importFirstDataColumnIndex"));
        }

        final String importCsvCharset = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getImportCsvCharset()))
                .filter(StringUtils::isNotBlank)
                .orElseGet(workbookMeta::getImportCsvCharset);

        final char importCsvDelimiter = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getImportCsvDelimiter()))
                .filter(delimiter -> delimiter != PxlConstants.UNSPECIFIED_IMPORT_CSV_DELIMITER)
                .orElseGet(workbookMeta::getImportCsvDelimiter);

        final List<PxlImportColumnOption> importColumnOptions = Optional.ofNullable(sheetOption)
                .map(option -> option.getImportColumnOptions())
                .orElseGet(ArrayList::new);

        final List<PxlImportColumnMeta> importColumnMetas = new ArrayList<>();

        return new PxlImportSheetMeta(
                workbookMeta,                   // workbookMeta
                null,                           // sheetField
                rowConcreteCollectionClass,     // rowCollectionClass
                rowClass,                       // rowClass
                sheetNames,                     // candidateSheetNames
                importEnabled,                  // importEnabled
                importOverrideSuperClassSheet,  // importOverrideSuperClassSheet
                importExcludeHiddenRows,        // importExcludeHiddenRows
                importExcludeHiddenColumns,     // importExcludeHiddenColumns
                importEachCellOfMergedRegion,   // importEachCellOfMergedRegion
                importHeaderRowIndex,            // importHeaderRowIndex
                importFirstDataRowIndex,        // importFirstDataRowIndex
                importLastDataRowIndex,         // importLastDataRowIndex
                importFirstDataColumnIndex,     // importFirstDataColumnIndex
                importLastDataColumnIndex,      // importLastDataColumnIndex
                importCsvCharset,               // importCsvCharset
                importCsvDelimiter,             // importCsvDelimiter
                importColumnOptions,            // importColumnOptions
                importColumnMetas               // importColumnMetas
        );
    }

    /**
     * Resolves the ordered candidate sheet names, preferring the sheet option's names, then the {@link PxlSheet}
     * annotation names, and finally the field name; each name is i18n-translated, whitespace-stripped and blank entries dropped.
     *
     * @param workbookMeta    the workbook metadata supplying the import resource bundle
     * @param sheetOption     the sheet option whose names take priority; may be {@code null}
     * @param sheetAnnotation the sheet annotation providing fallback names
     * @param sheetField      the sheet field whose name is the final fallback
     * @return the non-empty ordered list of candidate sheet names
     * @throws PxlNullPointerException if {@code workbookMeta}, {@code sheetAnnotation}, or {@code sheetField} is {@code null}
     */
    private static List<String> makeCandidateSheetNames(final PxlImportWorkbookMeta workbookMeta,
                                                        @Nullable final PxlImportSheetOption sheetOption,
                                                        final PxlSheet sheetAnnotation,
                                                        final Field sheetField)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");
        PxlAssertSupport.notNull(sheetAnnotation, "sheetAnnotation");
        PxlAssertSupport.notNull(sheetField, "sheetField");

        final ResourceBundle importResourceBundle = Optional.ofNullable(workbookMeta)
                .map(PxlImportWorkbookMeta::getImportResourceBundle)
                .orElse(null);

        List<String> candidateSheetNames;

        candidateSheetNames = PxlImportSheetOption.getImportSheetNames(sheetOption);
        if (PxlCollectionUtils.isNotEmpty(candidateSheetNames)) {
            candidateSheetNames = candidateSheetNames.stream()
                    .map(name -> PxlI18nContent.translate(importResourceBundle, name))
                    .map(StringUtils::deleteWhitespace)
                    .filter(StringUtils::isNotBlank)
                    .collect(Collectors.toList());
            if (PxlCollectionUtils.isNotEmpty(candidateSheetNames)) {
                return candidateSheetNames;
            }
        }

        // Get the value of the @PxlSheet annotation, i.e. the sheet name.
        candidateSheetNames = Arrays.stream(sheetAnnotation.name())
                .map(name -> PxlI18nContent.translate(importResourceBundle, name))
                .map(StringUtils::deleteWhitespace)
                .filter(StringUtils::isNotBlank)
                .collect(Collectors.toList());
        if (PxlCollectionUtils.isNotEmpty(candidateSheetNames)) {
            return candidateSheetNames;
        }

        candidateSheetNames = Arrays.asList(sheetField.getName());

        return candidateSheetNames;
    }

    /**
     * Appends the given column metadata to this sheet's column metadata list.
     *
     * @param importColumnMetas the column metadata to add
     */
    public void addImportColumnMetas(final List<PxlImportColumnMeta> importColumnMetas) {

        this.importColumnMetas.addAll(importColumnMetas);
    }

    /**
     * Appends a single column metadata to this sheet's column metadata list.
     *
     * @param importColumnMeta the column metadata to add
     */
    public void addImportColumnMeta(final PxlImportColumnMeta importColumnMeta) {

        this.importColumnMetas.add(importColumnMeta);
    }

}
