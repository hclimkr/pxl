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
import io.github.hclimkr.pxl.option.PxlExportColumnOption;
import io.github.hclimkr.pxl.option.PxlExportSheetOption;
import io.github.hclimkr.pxl.styler.PxlStyler;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import io.github.hclimkr.pxl.util.PxlMiscUtils;
import lombok.Getter;
import lombok.Setter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.util.WorkbookUtil;

import javax.validation.constraints.NotEmpty;
import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Excel sheet export metadata.
 */
@Getter
public final class PxlExportSheetMeta {

    private final PxlExportWorkbookMeta workbookMeta;

    private final Field sheetField;

    private final Class<? extends Collection<?>> rowCollectionClass;

    private final Class<?> rowClass;

    private List<String> candidateSheetNames;

    private final boolean exportEnabled;

    private final boolean exportSampleEnabled;

    private final boolean exportOverrideSuperClassSheet;

    private final float exportRowHeightInPoints;

    private final String exportOrder;

    private final Field exportGroupingField;

    private final int exportHeaderRowIndex;          // 1-based

    private final int exportFirstDataRowIndex;      // 1-based inclusive

    private final int exportLastDataRowIndex;       // 1-based inclusive

    private final int exportFirstDataColumnIndex;   // 1-based inclusive

    private final int exportLastDataColumnIndex;    // 1-based inclusive

    private final boolean exportIfNull;

    private final boolean exportIfEmpty;

    private final boolean exportColumnFilter;

    private final Class<? extends PxlStyler> exportSheetRequiredHeaderCellStyler;

    private final Class<? extends PxlStyler> exportSheetOptionalHeaderCellStyler;

    private final Class<? extends PxlStyler> exportSheetDataCellStyler;

    private final List<PxlExportColumnOption> exportColumnOptions;

    private final List<PxlExportColumnMeta> exportColumnMetas;

    private final boolean isRequired;

    @Setter
    private String actualExportSheetName;

    @Setter
    private int actualExportSheetIndex = -1;

    @Setter
    private int actualExportHeaderRowIndex = -1;        // 0-based

    @Setter
    private int actualExportOriginDataRowIndex = -1;    // 0-based inclusive

    @Setter
    private int actualExportBoundDataRowIndex = -1;     // 0-based exclusive

    @Setter
    private int actualExportOriginDataColumnIndex = -1; // 0-based inclusive

    @Setter
    private int actualExportBoundDataColumnIndex = -1;  // 0-based exclusive

    /**
     * Creates the resolved export metadata for one sheet, storing the merged option/annotation values.
     *
     * @param workbookMeta                        the enclosing workbook metadata
     * @param sheetField                          the {@link PxlSheet} Collection field to bind, or {@code null} for an ad-hoc sheet
     * @param rowCollectionClass                  the Collection type holding the rows
     * @param rowClass                            the row (element) class whose columns are bound
     * @param candidateSheetNames                 the ordered candidate sheet names; the first is written as the sheet name
     * @param exportEnabled                       whether this sheet is exported
     * @param exportSampleEnabled                 whether this sheet is included in the sample workbook
     * @param exportOverrideSuperClassSheet       whether this sheet overrides a same-named super-class sheet
     * @param exportRowHeightInPoints             the default data row height in points
     * @param exportOrder                         the sort key used to order sheets
     * @param exportGroupingField                 the field whose value splits rows into grouped sheets, or {@code null}
     * @param exportHeaderRowIndex                the 1-based header row index
     * @param exportFirstDataRowIndex             the 1-based inclusive first data row index
     * @param exportLastDataRowIndex              the 1-based inclusive last data row index
     * @param exportFirstDataColumnIndex          the 1-based inclusive first data column index
     * @param exportLastDataColumnIndex           the 1-based inclusive last data column index
     * @param exportIfNull                        whether a sheet is created when the row collection is {@code null}
     * @param exportIfEmpty                       whether a sheet is created when the row collection is empty
     * @param exportColumnFilter                  whether an auto-filter is applied
     * @param exportSheetRequiredHeaderCellStyler the header cell styler for required columns
     * @param exportSheetOptionalHeaderCellStyler the header cell styler for optional columns
     * @param exportSheetDataCellStyler           the data cell styler
     * @param exportColumnOptions                 the per-column export overrides
     * @param exportColumnMetas                   the (initially empty) column metadata list
     */
    private PxlExportSheetMeta(final PxlExportWorkbookMeta workbookMeta,
                               final Field sheetField,
                               final Class<? extends Collection<?>> rowCollectionClass,
                               final Class<?> rowClass,
                               final List<String> candidateSheetNames,
                               final boolean exportEnabled,
                               final boolean exportSampleEnabled,
                               final boolean exportOverrideSuperClassSheet,
                               final float exportRowHeightInPoints,
                               final String exportOrder,
                               final Field exportGroupingField,
                               final int exportHeaderRowIndex,
                               final int exportFirstDataRowIndex,
                               final int exportLastDataRowIndex,
                               final int exportFirstDataColumnIndex,
                               final int exportLastDataColumnIndex,
                               final boolean exportIfNull,
                               final boolean exportIfEmpty,
                               final boolean exportColumnFilter,
                               final Class<? extends PxlStyler> exportSheetRequiredHeaderCellStyler,
                               final Class<? extends PxlStyler> exportSheetOptionalHeaderCellStyler,
                               final Class<? extends PxlStyler> exportSheetDataCellStyler,
                               final List<PxlExportColumnOption> exportColumnOptions,
                               final List<PxlExportColumnMeta> exportColumnMetas) {

        this.workbookMeta = workbookMeta;

        this.sheetField = sheetField;
        this.rowCollectionClass = rowCollectionClass;
        this.rowClass = rowClass;
        this.candidateSheetNames = candidateSheetNames;
        this.exportEnabled = exportEnabled;
        this.exportSampleEnabled = exportSampleEnabled;
        this.exportOverrideSuperClassSheet = exportOverrideSuperClassSheet;
        this.exportRowHeightInPoints = exportRowHeightInPoints;
        this.exportOrder = exportOrder;
        this.exportGroupingField = exportGroupingField;
        this.exportHeaderRowIndex = exportHeaderRowIndex;
        this.exportFirstDataRowIndex = exportFirstDataRowIndex;
        this.exportLastDataRowIndex = exportLastDataRowIndex;
        this.exportFirstDataColumnIndex = exportFirstDataColumnIndex;
        this.exportLastDataColumnIndex = exportLastDataColumnIndex;
        this.exportIfNull = exportIfNull;
        this.exportIfEmpty = exportIfEmpty;
        this.exportColumnFilter = exportColumnFilter;
        this.exportSheetRequiredHeaderCellStyler = exportSheetRequiredHeaderCellStyler;
        this.exportSheetOptionalHeaderCellStyler = exportSheetOptionalHeaderCellStyler;
        this.exportSheetDataCellStyler = exportSheetDataCellStyler;
        this.exportColumnOptions = exportColumnOptions;
        this.exportColumnMetas = exportColumnMetas;

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
     * On export, collects the sheet metadata for the Excel file from the workbook class.
     * Each {@link PxlSheet}-annotated Collection field becomes one sheet; sheet options (matched by field name,
     * falling back to the wildcard) override the annotation, i18n-resolved names are computed, sheets are sorted by
     * {@code exportOrder}, and actual 0-based sheet indexes/names are assigned to the enabled sheets.
     *
     * @param workbookClass the annotated workbook class whose fields are scanned for sheets; may be {@code null}
     * @param workbookMeta  the enclosing workbook metadata, supplying cascaded styler defaults and the i18n bundle
     * @param sheetOptions  per-field runtime sheet overrides; may be {@code null}
     * @param isForSample   {@code true} to select sheets by {@code exportSampleEnabled}, {@code false} by {@code exportEnabled}
     * @return the collected sheet metadata list, sorted by export order
     * @throws PxlNullPointerException if {@code workbookMeta} is {@code null}
     * @throws PxlReflectionException  if a sheet field's parameterized type cannot be resolved
     * @throws PxlArgumentException    if the grouping field ({@code exportGroupingFieldName}) is not found on the row class
     * @throws PxlDataException        if a row/column index is negative or inconsistent, or exported sheet names collide
     */
    public static List<PxlExportSheetMeta> makeExportSheetMetas(@Nullable final Class<?> workbookClass,
                                                                final PxlExportWorkbookMeta workbookMeta,
                                                                @Nullable final List<PxlExportSheetOption> sheetOptions,
                                                                final boolean isForSample)
            throws PxlNullPointerException, PxlReflectionException, PxlArgumentException, PxlDataException {

        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");

//        final Field[] sheetFields = workbookClass.getDeclaredFields();
        final List<Field> sheetFields = PxlReflectionSupport.getAllFields(workbookClass);
        final List<PxlExportSheetMeta> sheetMetas = new ArrayList<>(PxlCollectionUtils.size(sheetFields));
        final Set<String> overriddenSheetNames = new HashSet<>();

        for (final Field sheetField : sheetFields) {
            // Get the @PxlSheet annotation for the sheet field.
            final PxlSheet sheetAnnotation = sheetField.getAnnotation(PxlSheet.class);
            if (Objects.isNull(sheetAnnotation)) {
                continue;
            }

            // Ignore if it is not a Collection type.
            if (!PxlClassSupport.isCollectionClass(sheetField.getType())) {
                continue;
            }

            @SuppressWarnings("unchecked") final Class<? extends Collection<?>> rowCollectionClass =
                    (Class<? extends Collection<?>>) sheetField.getType();

            // Get the generic class of the Collection.
            final Class<?> rowClass = PxlReflectionSupport.getParameterizedArgument0(sheetField);

            final PxlExportSheetOption sheetOption = Optional.ofNullable(sheetOptions)
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

            final boolean exportEnabled = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getExportEnabled()))
                    .orElseGet(sheetAnnotation::exportEnabled);

            final boolean exportSampleEnabled = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getExportSampleEnabled()))
                    .orElseGet(sheetAnnotation::exportSampleEnabled);

            final boolean exportOverrideSuperClassSheet = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getExportOverrideSuperClassSheet()))
                    .orElseGet(sheetAnnotation::exportOverrideSuperClassSheet);

            final float exportRowHeightInPoints = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getExportRowHeightInPoints()))
                    .orElseGet(sheetAnnotation::exportRowHeightInPoints);

            final String exportOrder = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getExportOrder()))
                    .orElseGet(sheetAnnotation::exportOrder);

            final String exportGroupingFieldName = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getExportGroupingFieldName()))
                    .orElseGet(sheetAnnotation::exportGroupingFieldName);
            Field exportGroupingField = null;
            if (StringUtils.isNotBlank(exportGroupingFieldName)) {
                exportGroupingField = PxlReflectionSupport.getAllFields(rowClass).stream()
                        .filter(f -> f.getName().equals(exportGroupingFieldName))
                        .findFirst()
                        .orElse(null);
                if (Objects.isNull(exportGroupingField)) {
                    throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_GROUPING_FIELD_NOT_FOUND, exportGroupingFieldName, rowClass.getSimpleName()));
                }
            }

            final int exportHeaderRowIndex = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getExportHeaderRowIndex()))
                    .orElseGet(sheetAnnotation::exportHeaderRowIndex);
            if (exportHeaderRowIndex < 0) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "exportHeaderRowIndex"));
            }

            final int exportFirstDataRowIndex = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getExportFirstDataRowIndex()))
                    .orElseGet(sheetAnnotation::exportFirstDataRowIndex);
            if (exportFirstDataRowIndex < 0) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "exportFirstDataRowIndex"));
            }

            final int exportLastDataRowIndex = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getExportLastDataRowIndex()))
                    .orElseGet(sheetAnnotation::exportLastDataRowIndex);
            if (exportLastDataRowIndex < 0) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "exportLastDataRowIndex"));
            }

            final int exportFirstDataColumnIndex = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getExportFirstDataColumnIndex()))
                    .orElseGet(sheetAnnotation::exportFirstDataColumnIndex);
            if (exportFirstDataColumnIndex < 0) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "exportFirstDataColumnIndex"));
            }

            final int exportLastDataColumnIndex = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getExportLastDataColumnIndex()))
                    .orElseGet(sheetAnnotation::exportLastDataColumnIndex);
            if (exportLastDataColumnIndex < 0) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "exportLastDataColumnIndex"));
            }

            if (exportHeaderRowIndex != PxlConstants.DEFAULT_EXPORT_HEADER_ROW_INDEX
                    && exportFirstDataRowIndex != PxlConstants.DEFAULT_EXPORT_FIRST_DATA_ROW_INDEX
                    && exportFirstDataRowIndex <= exportHeaderRowIndex) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_FIRST_DATA_ROW_GT_HEADER, candidateSheetNames, "exportFirstDataRowIndex", "exportHeaderRowIndex"));
            }

            if (exportFirstDataRowIndex != PxlConstants.DEFAULT_EXPORT_FIRST_DATA_ROW_INDEX
                    && exportLastDataRowIndex != PxlConstants.DEFAULT_EXPORT_LAST_DATA_ROW_INDEX
                    && exportLastDataRowIndex < exportFirstDataRowIndex) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_LAST_DATA_GE_FIRST_DATA, candidateSheetNames, "exportLastDataRowIndex", "exportFirstDataRowIndex"));
            }

            if (exportFirstDataColumnIndex != PxlConstants.DEFAULT_EXPORT_FIRST_DATA_COLUMN_INDEX
                    && exportLastDataColumnIndex != PxlConstants.DEFAULT_EXPORT_LAST_DATA_COLUMN_INDEX
                    && exportLastDataColumnIndex < exportFirstDataColumnIndex) {
                throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_LAST_DATA_GE_FIRST_DATA, candidateSheetNames, "exportLastDataColumnIndex", "exportFirstDataColumnIndex"));
            }

            final boolean exportIfNull = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getExportIfNull()))
                    .orElseGet(sheetAnnotation::exportIfNull);
            final boolean exportIfEmpty = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getExportIfEmpty()))
                    .orElseGet(sheetAnnotation::exportIfEmpty);
            final boolean exportColumnFilter = Optional.ofNullable(sheetOption)
                    .flatMap(option -> Optional.ofNullable(option.getExportColumnFilter()))
                    .orElseGet(sheetAnnotation::exportColumnFilter);

            Class<? extends PxlStyler> exportSheetRequiredHeaderCellStyler =
                    Objects.nonNull(sheetOption) && Objects.nonNull(sheetOption.getExportSheetRequiredHeaderCellStyler())
                            ? sheetOption.getExportSheetRequiredHeaderCellStyler()
                            : sheetAnnotation.exportSheetRequiredHeaderCellStyler();
            if (!PxlMiscUtils.isEffectiveCellStylerClass(exportSheetRequiredHeaderCellStyler)) {
                exportSheetRequiredHeaderCellStyler = workbookMeta.getExportWorkbookRequiredHeaderCellStyler();
            }

            Class<? extends PxlStyler> exportSheetOptionalHeaderCellStyler =
                    Objects.nonNull(sheetOption) && Objects.nonNull(sheetOption.getExportSheetOptionalHeaderCellStyler())
                            ? sheetOption.getExportSheetOptionalHeaderCellStyler()
                            : sheetAnnotation.exportSheetOptionalHeaderCellStyler();
            if (!PxlMiscUtils.isEffectiveCellStylerClass(exportSheetOptionalHeaderCellStyler)) {
                exportSheetOptionalHeaderCellStyler = workbookMeta.getExportWorkbookOptionalHeaderCellStyler();
            }

            Class<? extends PxlStyler> exportSheetDataCellStyler =
                    Objects.nonNull(sheetOption) && Objects.nonNull(sheetOption.getExportSheetDataCellStyler())
                            ? sheetOption.getExportSheetDataCellStyler()
                            : sheetAnnotation.exportSheetDataCellStyler();
            if (!PxlMiscUtils.isEffectiveCellStylerClass(exportSheetDataCellStyler)) {
                exportSheetDataCellStyler = workbookMeta.getExportWorkbookDataCellStyler();
            }

            final List<PxlExportColumnOption> exportColumnOptions = Optional.ofNullable(sheetOption)
                    .map(option -> option.getExportColumnOptions())
                    .orElseGet(ArrayList::new);

            final List<PxlExportColumnMeta> exportColumnMetas = new ArrayList<>();

            sheetField.setAccessible(true);
            sheetMetas.add(
                    new PxlExportSheetMeta(
                            workbookMeta,                   // workbookMeta
                            sheetField,                     // sheetField
                            rowCollectionClass,             // rowCollectionClass
                            rowClass,                       // rowClass
                            candidateSheetNames,            // candidateSheetNames
                            exportEnabled,                  // exportEnabled
                            exportSampleEnabled,            // exportSampleEnabled
                            exportOverrideSuperClassSheet,  // exportOverrideSuperClassSheet
                            exportRowHeightInPoints,        // exportRowHeightInPoints
                            exportOrder,                    // exportOrder
                            exportGroupingField,            // exportGroupingField
                            exportHeaderRowIndex,            // exportHeaderRowIndex
                            exportFirstDataRowIndex,        // exportFirstDataRowIndex
                            exportLastDataRowIndex,         // exportLastDataRowIndex
                            exportFirstDataColumnIndex,     // exportFirstDataColumnIndex
                            exportLastDataColumnIndex,      // exportLastDataColumnIndex
                            exportIfNull,                   // exportIfNull
                            exportIfEmpty,                  // exportIfEmpty
                            exportColumnFilter,             // exportColumnFilter
                            exportSheetRequiredHeaderCellStyler,  // exportSheetRequiredHeaderCellStyler
                            exportSheetOptionalHeaderCellStyler,   // exportSheetOptionalHeaderCellStyler
                            exportSheetDataCellStyler,            // exportSheetDataCellStyler
                            exportColumnOptions,            // exportColumnOptions
                            exportColumnMetas               // exportColumnMetas
                    ));

            if (exportEnabled && exportOverrideSuperClassSheet) {
                overriddenSheetNames.addAll(candidateSheetNames);
            }
        }

        sheetMetas.sort(Comparator.comparing(PxlExportSheetMeta::getExportOrder));

        int exportSheetIndex = 0;
        for (final PxlExportSheetMeta sheetMeta : sheetMetas) {
            if (!isForSample && !sheetMeta.isExportEnabled()) {
                continue;
            }
            if (isForSample && !sheetMeta.isExportSampleEnabled()) {
                continue;
            }

            sheetMeta.setActualExportSheetIndex(exportSheetIndex);
            // export using the first sheet name.
            sheetMeta.setActualExportSheetName(WorkbookUtil.createSafeSheetName(PxlCollectionUtils.get(sheetMeta.getCandidateSheetNames(), 0)));

            exportSheetIndex++;
        }

        final List<String> exportedSheetNames = sheetMetas.stream()
                .map(PxlExportSheetMeta::getActualExportSheetName)
                .filter(Objects::nonNull)
                .collect(Collectors.toList());
        final Set<String> duplicatedSheetNames = PxlCollectionUtils.findDuplicates(exportedSheetNames);
        if (!duplicatedSheetNames.isEmpty()) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_DUPLICATE_SHEET_NAME, duplicatedSheetNames));
        }

        return sheetMetas;
    }

    /**
     * On export, collects the sheet metadata for the Excel file from the method parameters.
     * Used when exporting an ad-hoc single sheet: the sheet name is trimmed and used as the sole candidate name,
     * and the sheet option (when present) overrides the built-in defaults.
     *
     * @param sheetName          the target sheet name; must not be blank
     * @param rowCollectionClass the Collection type holding the rows; may be {@code null}, must be a Collection type when given
     * @param rowClass           the row (element) class whose columns are bound
     * @param workbookMeta       the enclosing workbook metadata, supplying cascaded styler defaults
     * @param sheetOption        runtime overrides for this sheet; may be {@code null}
     * @return the single sheet metadata, with actual sheet index 0 and its safe name assigned
     * @throws PxlNullPointerException if {@code sheetName}, {@code rowClass}, or {@code workbookMeta} is {@code null}
     * @throws PxlArgumentException    if {@code sheetName} is blank, or the grouping field ({@code exportGroupingFieldName}) is not found on the row class
     * @throws PxlDataException        if {@code rowCollectionClass} is not a Collection type, or a row/column index is negative or inconsistent
     */
    public static PxlExportSheetMeta makeExportSheetMeta(final String sheetName,
                                                         @Nullable final Class<?> rowCollectionClass,
                                                         final Class<?> rowClass,
                                                         final PxlExportWorkbookMeta workbookMeta,
                                                         @Nullable final PxlExportSheetOption sheetOption)
            throws PxlNullPointerException, PxlArgumentException, PxlDataException {

        PxlAssertSupport.notBlank(sheetName, "sheetName");
        PxlAssertSupport.notNull(rowClass, "rowClass");
        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");

        if (Objects.nonNull(rowCollectionClass) && !PxlClassSupport.isCollectionClass(rowCollectionClass)) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_NOT_COLLECTION_TYPE, rowCollectionClass.getSimpleName()));
        }

        @SuppressWarnings("unchecked") final Class<? extends Collection<?>> rowCollectionClassTyped =
                (Class<? extends Collection<?>>) rowCollectionClass;

        if (StringUtils.isBlank(sheetName)) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_NAME_INVALID, sheetName));
        }

        final List<String> candidateSheetNames = Collections.singletonList(StringUtils.trim(sheetName));

        final boolean exportEnabled = true;
        final boolean exportSampleEnabled = true;

        final boolean exportOverrideSuperClassSheet = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getExportOverrideSuperClassSheet()))
                .orElse(PxlConstants.DEFAULT_EXPORT_OVERRIDE_SUPER_CLASS_SHEET);

        final float exportRowHeightInPoints = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getExportRowHeightInPoints()))
                .orElse(PxlConstants.DEFAULT_EXPORT_ROW_HEIGHT_IN_POINTS);

        final String exportOrder = "";
        final String exportGroupingFieldName = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getExportGroupingFieldName()))
                .orElse(null);
        Field exportGroupingField = null;
        if (StringUtils.isNotBlank(exportGroupingFieldName)) {
            exportGroupingField = PxlReflectionSupport.getAllFields(rowClass).stream()
                    .filter(f -> f.getName().equals(exportGroupingFieldName))
                    .findFirst()
                    .orElse(null);
            if (Objects.isNull(exportGroupingField)) {
                throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_GROUPING_FIELD_NOT_FOUND, exportGroupingFieldName, rowClass.getSimpleName()));
            }
        }

        final int exportHeaderRowIndex = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getExportHeaderRowIndex()))
                .orElse(PxlConstants.DEFAULT_EXPORT_HEADER_ROW_INDEX);
        if (exportHeaderRowIndex < 0) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "exportHeaderRowIndex"));
        }

        final int exportFirstDataRowIndex = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getExportFirstDataRowIndex()))
                .orElse(PxlConstants.DEFAULT_EXPORT_FIRST_DATA_ROW_INDEX);
        if (exportFirstDataRowIndex < 0) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "exportFirstDataRowIndex"));
        }

        final int exportLastDataRowIndex = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getExportLastDataRowIndex()))
                .orElse(PxlConstants.DEFAULT_EXPORT_LAST_DATA_ROW_INDEX);
        if (exportLastDataRowIndex < 0) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "exportLastDataRowIndex"));
        }

        final int exportFirstDataColumnIndex = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getExportFirstDataColumnIndex()))
                .orElse(PxlConstants.DEFAULT_EXPORT_FIRST_DATA_COLUMN_INDEX);
        if (exportFirstDataColumnIndex < 0) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "exportFirstDataColumnIndex"));
        }

        final int exportLastDataColumnIndex = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getExportLastDataColumnIndex()))
                .orElse(PxlConstants.DEFAULT_EXPORT_LAST_DATA_COLUMN_INDEX);
        if (exportLastDataColumnIndex < 0) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_INDEX_NON_NEGATIVE, candidateSheetNames, "exportLastDataColumnIndex"));
        }

        if (exportHeaderRowIndex != PxlConstants.DEFAULT_EXPORT_HEADER_ROW_INDEX
                && exportFirstDataRowIndex != PxlConstants.DEFAULT_EXPORT_FIRST_DATA_ROW_INDEX
                && exportFirstDataRowIndex <= exportHeaderRowIndex) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_FIRST_DATA_ROW_GT_HEADER, candidateSheetNames, "exportFirstDataRowIndex", "exportHeaderRowIndex"));
        }

        if (exportFirstDataRowIndex != PxlConstants.DEFAULT_EXPORT_FIRST_DATA_ROW_INDEX
                && exportLastDataRowIndex != PxlConstants.DEFAULT_EXPORT_LAST_DATA_ROW_INDEX
                && exportLastDataRowIndex < exportFirstDataRowIndex) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_LAST_DATA_GE_FIRST_DATA, candidateSheetNames, "exportLastDataRowIndex", "exportFirstDataRowIndex"));
        }

        if (exportFirstDataColumnIndex != PxlConstants.DEFAULT_EXPORT_FIRST_DATA_COLUMN_INDEX
                && exportLastDataColumnIndex != PxlConstants.DEFAULT_EXPORT_LAST_DATA_COLUMN_INDEX
                && exportLastDataColumnIndex < exportFirstDataColumnIndex) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_SHEET_LAST_DATA_GE_FIRST_DATA, candidateSheetNames, "exportLastDataColumnIndex", "exportFirstDataColumnIndex"));
        }

        final boolean exportIfNull = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getExportIfNull()))
                .orElse(PxlConstants.DEFAULT_EXPORT_IF_NULL);
        final boolean exportIfEmpty = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getExportIfEmpty()))
                .orElse(PxlConstants.DEFAULT_EXPORT_IF_EMPTY);
        final boolean exportColumnFilter = Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getExportColumnFilter()))
                .orElse(PxlConstants.DEFAULT_EXPORT_COLUMN_FILTER);

        Class<? extends PxlStyler> exportSheetRequiredHeaderCellStyler =
                Objects.nonNull(sheetOption) && Objects.nonNull(sheetOption.getExportSheetRequiredHeaderCellStyler())
                        ? sheetOption.getExportSheetRequiredHeaderCellStyler()
                        : null;
        if (!PxlMiscUtils.isEffectiveCellStylerClass(exportSheetRequiredHeaderCellStyler)) {
            exportSheetRequiredHeaderCellStyler = workbookMeta.getExportWorkbookRequiredHeaderCellStyler();
        }

        Class<? extends PxlStyler> exportSheetOptionalHeaderCellStyler =
                Objects.nonNull(sheetOption) && Objects.nonNull(sheetOption.getExportSheetOptionalHeaderCellStyler())
                        ? sheetOption.getExportSheetOptionalHeaderCellStyler()
                        : null;
        if (!PxlMiscUtils.isEffectiveCellStylerClass(exportSheetOptionalHeaderCellStyler)) {
            exportSheetOptionalHeaderCellStyler = workbookMeta.getExportWorkbookOptionalHeaderCellStyler();
        }

        Class<? extends PxlStyler> exportSheetDataCellStyler =
                Objects.nonNull(sheetOption) && Objects.nonNull(sheetOption.getExportSheetDataCellStyler())
                        ? sheetOption.getExportSheetDataCellStyler()
                        : null;
        if (!PxlMiscUtils.isEffectiveCellStylerClass(exportSheetDataCellStyler)) {
            exportSheetDataCellStyler = workbookMeta.getExportWorkbookDataCellStyler();
        }

        final List<PxlExportColumnOption> exportColumnOptions = Optional.ofNullable(sheetOption)
                .map(option -> option.getExportColumnOptions())
                .orElseGet(ArrayList::new);

        final List<PxlExportColumnMeta> exportColumnMetas = new ArrayList<>();

        final PxlExportSheetMeta sheetMeta = new PxlExportSheetMeta(
                workbookMeta,                   // workbookMeta
                null,                           // sheetField
                rowCollectionClassTyped,        // rowCollectionClass
                rowClass,                       // rowClass
                candidateSheetNames,            // candidateSheetNames
                exportEnabled,                  // exportEnabled
                exportSampleEnabled,            // exportSampleEnabled
                exportOverrideSuperClassSheet,  // exportOverrideSuperClassSheet
                exportRowHeightInPoints,        // exportRowHeightInPoints
                exportOrder,                    // exportOrder
                exportGroupingField,            // exportGroupingField
                exportHeaderRowIndex,            // exportHeaderRowIndex
                exportFirstDataRowIndex,        // exportFirstDataRowIndex
                exportLastDataRowIndex,         // exportLastDataRowIndex
                exportFirstDataColumnIndex,     // exportFirstDataColumnIndex
                exportLastDataColumnIndex,      // exportLastDataColumnIndex
                exportIfNull,                   // exportIfNull
                exportIfEmpty,                  // exportIfEmpty
                exportColumnFilter,             // exportColumnFilter
                exportSheetRequiredHeaderCellStyler,  // exportSheetRequiredHeaderCellStyler
                exportSheetOptionalHeaderCellStyler,   // exportSheetOptionalHeaderCellStyler
                exportSheetDataCellStyler,            // exportSheetDataCellStyler
                exportColumnOptions,            // exportColumnOptions
                exportColumnMetas               // exportColumnMetas
        );

        sheetMeta.setActualExportSheetIndex(0);
        // export using the first sheet name.
        sheetMeta.setActualExportSheetName(WorkbookUtil.createSafeSheetName(PxlCollectionUtils.get(sheetMeta.getCandidateSheetNames(), 0)));

        return sheetMeta;
    }

    /**
     * Resolves the ordered candidate sheet names, preferring the sheet option's names, then the {@link PxlSheet}
     * annotation names, and finally the field name; each name is i18n-translated and blank entries are dropped.
     *
     * @param workbookMeta    the workbook metadata supplying the export resource bundle
     * @param sheetOption     the sheet option whose names take priority; may be {@code null}
     * @param sheetAnnotation the sheet annotation providing fallback names
     * @param sheetField      the sheet field whose name is the final fallback
     * @return the non-empty ordered list of candidate sheet names
     * @throws PxlNullPointerException if {@code workbookMeta}, {@code sheetAnnotation}, or {@code sheetField} is {@code null}
     */
    private static List<String> makeCandidateSheetNames(final PxlExportWorkbookMeta workbookMeta,
                                                        @Nullable final PxlExportSheetOption sheetOption,
                                                        final PxlSheet sheetAnnotation,
                                                        final Field sheetField)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");
        PxlAssertSupport.notNull(sheetAnnotation, "sheetAnnotation");
        PxlAssertSupport.notNull(sheetField, "sheetField");

        final ResourceBundle exportResourceBundle = Optional.ofNullable(workbookMeta)
                .map(PxlExportWorkbookMeta::getExportResourceBundle)
                .orElse(null);

        List<String> candidateSheetNames;

        candidateSheetNames = PxlExportSheetOption.getExportSheetNames(sheetOption);
        if (PxlCollectionUtils.isNotEmpty(candidateSheetNames)) {
            candidateSheetNames = candidateSheetNames.stream()
                    .map(name -> PxlI18nContent.translate(exportResourceBundle, name))
                    // .map(StringUtils::trim)
                    .filter(StringUtils::isNotBlank)
                    .collect(Collectors.toList());
            if (PxlCollectionUtils.isNotEmpty(candidateSheetNames)) {
                return candidateSheetNames;
            }
        }

        // Get the value of the @PxlSheet annotation, i.e. the sheet name.
        candidateSheetNames = Arrays.stream(sheetAnnotation.name())
                .map(name -> PxlI18nContent.translate(exportResourceBundle, name))
                .map(StringUtils::trim)
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
     * @param exportColumnMetas the column metadata to add
     */
    public void addExportColumnMetas(final List<PxlExportColumnMeta> exportColumnMetas) {

        this.exportColumnMetas.addAll(exportColumnMetas);
    }

    /**
     * Appends a single column metadata to this sheet's column metadata list.
     *
     * @param exportColumnMeta the column metadata to add
     */
    public void addExportColumnMeta(final PxlExportColumnMeta exportColumnMeta) {

        this.exportColumnMetas.add(exportColumnMeta);
    }

    /**
     * Returns whether this sheet has at least one column exported as a formula.
     *
     * @return {@code true} if any column is exported as a formula
     */
    public boolean hasAnyExportStringAsFormulaColumn() {

        return PxlCollectionUtils.emptyIfNull(this.exportColumnMetas)
                .stream()
                .anyMatch(PxlExportColumnMeta::isExportStringAsFormula);
    }

}
