package io.github.hclimkr.pxl.internal.meta;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.annotation.PxlExportConverter;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlDataException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.exception.PxlReflectionException;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nContent;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.support.*;
import io.github.hclimkr.pxl.option.PxlExportColumnOption;
import io.github.hclimkr.pxl.styler.PxlStyler;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import io.github.hclimkr.pxl.util.PxlMiscUtils;
import lombok.Getter;
import lombok.Setter;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;

import javax.validation.constraints.NotBlank;
import javax.validation.constraints.NotEmpty;
import javax.validation.constraints.NotNull;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * Excel column export metadata.
 */
@Getter
public final class PxlExportColumnMeta {

    private final PxlExportWorkbookMeta workbookMeta;

    private final PxlExportSheetMeta sheetMeta;

    private final Field columnField;

    private final Class<?> columnClass;

    private final List<String> candidateColumnNames;

    private final boolean exportEnabled;

    private final boolean exportSampleEnabled;

    private final String exportSample;

    private final boolean exportTrim;

    private final String exportPattern;

    private final int exportColumnWidth;

    private final String exportCollectionSeparator;

    private final boolean exportOverrideSuperClassColumn;

    private final String exportOrder;

    private final String exportMasking;

    private final Pattern exportMaskingPattern;

    private final String[] exportOptionItems;

    private final PxlColumn.ExportEnumDropDownListStyle exportEnumDropDownListStyle;

    private final String exportNullString;

    private final String exportTrueString;

    private final String exportFalseString;

    private final boolean exportStringAsPicture;

    private final boolean exportStringAsFormula;

    private final Class<? extends PxlStyler> exportColumnRequiredHeaderCellStyler;

    private final Class<? extends PxlStyler> exportColumnOptionalHeaderCellStyler;

    private final Class<? extends PxlStyler> exportColumnDataCellStyler;

    private final boolean isRequired;

    private DecimalFormat exportDecimalFormatterCache = null;

    private SimpleDateFormat exportJavaDateFormatterCache = null;

    private DateTimeFormatter exportDateTimeFormatterCache = null;

    private PxlExportConverterMeta exportCustomConverterMeta = null;

    @Setter
    private String actualExportColumnName;

    @Setter
    private int actualExportColumnIndex = -1;

    /**
     * Creates the resolved export metadata for one column, storing the merged option/annotation values,
     * compiling the masking pattern and, when export (or sample) is enabled, pre-compiling the export
     * number/date pattern and custom converter.
     *
     * @param workbookMeta                         the enclosing workbook metadata
     * @param sheetMeta                            the enclosing sheet metadata
     * @param columnField                          the column field to bind
     * @param candidateColumnNames                 the ordered candidate column names; the first is written as the header
     * @param exportEnabled                        whether this column is exported
     * @param exportSampleEnabled                  whether this column is included in the sample sheet
     * @param exportSample                         the sample cell value
     * @param exportTrim                           whether string values are trimmed on export
     * @param exportPattern                        the number/date format pattern; blank for none
     * @param exportColumnWidth                    the column width setting
     * @param exportCollectionSeparator            the element separator used when writing Collection columns
     * @param exportOverrideSuperClassColumn       whether this column overrides a same-named super-class column
     * @param exportOrder                          the sort key used to order columns
     * @param exportMasking                        the masking regular expression; blank for none
     * @param exportOptionItems                    the explicit dropdown option items
     * @param exportEnumDropDownListStyle          the dropdown style for enum columns
     * @param exportNullString                     the string written for {@code null} values
     * @param exportTrueString                     the string written for boolean {@code true} (kept only for Boolean columns)
     * @param exportFalseString                    the string written for boolean {@code false} (kept only for Boolean columns)
     * @param exportStringAsPicture                whether string values are exported as pictures
     * @param exportStringAsFormula                whether string values are exported as formulas
     * @param exportColumnRequiredHeaderCellStyler the header cell styler for required columns
     * @param exportColumnOptionalHeaderCellStyler the header cell styler for optional columns
     * @param exportColumnDataCellStyler           the data cell styler
     * @throws PxlArgumentException   if the masking or export pattern is malformed, or a custom converter has an invalid signature
     * @throws PxlReflectionException if the Collection element type cannot be resolved
     */
    private PxlExportColumnMeta(final PxlExportWorkbookMeta workbookMeta,
                                final PxlExportSheetMeta sheetMeta,
                                final Field columnField,
                                final List<String> candidateColumnNames,
                                final boolean exportEnabled,
                                final boolean exportSampleEnabled,
                                final String exportSample,
                                final boolean exportTrim,
                                final String exportPattern,
                                final int exportColumnWidth,
                                final String exportCollectionSeparator,
                                final boolean exportOverrideSuperClassColumn,
                                final String exportOrder,
                                final String exportMasking,
                                final String[] exportOptionItems,
                                final PxlColumn.ExportEnumDropDownListStyle exportEnumDropDownListStyle,
                                final String exportNullString,
                                final String exportTrueString,
                                final String exportFalseString,
                                final boolean exportStringAsPicture,
                                final boolean exportStringAsFormula,
                                final Class<? extends PxlStyler> exportColumnRequiredHeaderCellStyler,
                                final Class<? extends PxlStyler> exportColumnOptionalHeaderCellStyler,
                                final Class<? extends PxlStyler> exportColumnDataCellStyler)
            throws PxlArgumentException, PxlReflectionException {

        this.workbookMeta = workbookMeta;
        this.sheetMeta = sheetMeta;

        this.columnField = columnField;
        this.columnClass = columnField.getType();

        this.candidateColumnNames = candidateColumnNames;
        this.exportEnabled = exportEnabled;
        this.exportSampleEnabled = exportSampleEnabled;
        this.exportSample = exportSample;
        this.exportTrim = exportTrim;
        this.exportPattern = exportPattern;
        this.exportColumnWidth = exportColumnWidth;
        this.exportCollectionSeparator = exportCollectionSeparator;
        this.exportOverrideSuperClassColumn = exportOverrideSuperClassColumn;
        this.exportOrder = exportOrder;
        this.exportMasking = exportMasking;
        if (StringUtils.isNotBlank(this.exportMasking)) {
            try {
                this.exportMaskingPattern = Pattern.compile(this.exportMasking);
            } catch (IllegalArgumentException illegalArgumentException) {
                throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_MASKING_INVALID, this.exportMasking), illegalArgumentException);
            }
        } else {
            this.exportMaskingPattern = null;
        }
        this.exportOptionItems = exportOptionItems;
        this.exportEnumDropDownListStyle = exportEnumDropDownListStyle;
        this.exportNullString = exportNullString;
        this.exportTrueString = exportTrueString;
        this.exportFalseString = exportFalseString;
        this.exportStringAsPicture = exportStringAsPicture;
        this.exportStringAsFormula = exportStringAsFormula;
        this.exportColumnRequiredHeaderCellStyler = exportColumnRequiredHeaderCellStyler;
        this.exportColumnOptionalHeaderCellStyler = exportColumnOptionalHeaderCellStyler;
        this.exportColumnDataCellStyler = exportColumnDataCellStyler;

        this.isRequired = (Objects.nonNull(columnField.getAnnotation(NotNull.class)))
                || (Objects.nonNull(columnField.getAnnotation(NotEmpty.class)))
                || (Objects.nonNull(columnField.getAnnotation(NotBlank.class)));

        if (exportEnabled || exportSampleEnabled) {
            if (StringUtils.isNotBlank(exportPattern)) {
                // For a Collection column, the element type is the target of the pattern.
                final Class<?> patternTargetClass = PxlClassSupport.isCollectionClass(this.columnClass)
                        ? PxlReflectionSupport.getParameterizedArgument0(this.columnField)
                        : this.columnClass;

                if (PxlClassSupport.isNumberClass(patternTargetClass)) {
                    try {
                        this.exportDecimalFormatterCache = PxlNumberSupport.getDecimalFormat(exportPattern);
                    } catch (IllegalArgumentException illegalArgumentException) {
                        throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_EXPORT_NUMBER_PATTERN_INVALID, exportPattern), illegalArgumentException);
                    }
                } else if (PxlClassSupport.isJavaDateClass(patternTargetClass)) {
                    try {
                        this.exportJavaDateFormatterCache = PxlDateTimeSupport.getCellSimpleDateFormatter(exportPattern, Locale.ROOT);
                    } catch (IllegalArgumentException illegalArgumentException) {
                        throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_EXPORT_DATE_TIME_PATTERN_INVALID, exportPattern), illegalArgumentException);
                    }
                } else if (PxlClassSupport.isDateTimeClass(patternTargetClass)) {
                    try {
                        this.exportDateTimeFormatterCache = PxlDateTimeSupport.getCellDateTimeFormatter(exportPattern, Locale.ROOT);
                    } catch (IllegalArgumentException illegalArgumentException) {
                        throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_EXPORT_DATE_TIME_PATTERN_INVALID, exportPattern), illegalArgumentException);
                    }
                }
            }

            final Class<?> targetClass;
            if (PxlClassSupport.isCollectionClass(this.columnClass)) {
                targetClass = PxlReflectionSupport.getParameterizedArgument0(this.columnField);
            } else {
                targetClass = this.columnClass;
            }

            if (PxlClassSupport.isCustomConvertableClass(targetClass)) {
                this.exportCustomConverterMeta = PxlExportConverterMeta.of(targetClass);
            }
        }
    }

    /**
     * Returns a debug string of the candidate column names and the bound field name.
     *
     * @return the string representation
     */
    @Override
    public String toString() {

        return "[" + candidateColumnNames + ", " + columnField.getName() + "]";
    }

    /**
     * Returns whether this column uses a custom export converter ({@link PxlExportConverter} or {@code toString}).
     *
     * @return {@code true} if a custom export converter is resolved for this column
     */
    public boolean isExportCustomConvertable() {

        return Objects.nonNull(exportCustomConverterMeta);
    }

    /**
     * On export, collects the sheet's column metadata from the row class.
     * Each {@link PxlColumn}-annotated field becomes one column; column options (matched by field name) override the
     * annotation, i18n-resolved names are computed, columns are sorted by {@code exportOrder}, and actual 0-based
     * column indexes/names are assigned within the sheet's configured column range.
     * <p>
     * Besides the names, two more values are read through the workbook's export bundle. The {@code exportSample} of a
     * String or enum column is translated, element by element when the column is a Collection; and the
     * {@code exportOptionItems} of a String column are translated as well, so the dropdown lists the same text the
     * sample cell holds. An enum, numeric or temporal column always writes its value in canonical form, so its
     * dropdown is left untranslated to stay consistent with what is written.
     *
     * @param sheetMeta   the enclosing sheet metadata, supplying the row class, column options and cascaded stylers
     * @param isForSample {@code true} to select columns by {@code exportSampleEnabled}, {@code false} by {@code exportEnabled}
     * @return the collected column metadata list, sorted by export order
     * @throws PxlNullPointerException if {@code sheetMeta} is {@code null}
     * @throws PxlReflectionException  if a column field's type cannot be resolved
     * @throws PxlArgumentException    if a column converter or styler declaration is invalid
     * @throws PxlDataException        if the column count exceeds the format limit, the configured column range drops columns, or exported column names collide
     */
    public static List<PxlExportColumnMeta> makeExportColumnMetas(final PxlExportSheetMeta sheetMeta,
                                                                  final boolean isForSample)
            throws PxlNullPointerException, PxlReflectionException, PxlArgumentException, PxlDataException {

        PxlAssertSupport.notNull(sheetMeta, "sheetMeta");

        final PxlExportWorkbookMeta workbookMeta = sheetMeta.getWorkbookMeta();
        final List<PxlExportColumnOption> columnOptions = sheetMeta.getExportColumnOptions();

        final ResourceBundle exportResourceBundle = Optional.ofNullable(workbookMeta)
                .map(PxlExportWorkbookMeta::getExportResourceBundle)
                .orElse(null);

        final Class<?> rowClass = sheetMeta.getRowClass();
//        final Field[] columnFields = rowClass.getDeclaredFields();
        final List<Field> columnFields = PxlReflectionSupport.getAllFields(rowClass);
        final List<PxlExportColumnMeta> columnMetas = new ArrayList<>(PxlCollectionUtils.size(columnFields));
        final Set<String> overriddenColumnNames = new HashSet<>();

        for (final Field columnField : columnFields) {
            final PxlColumn columnAnnotation = columnField.getAnnotation(PxlColumn.class);

            if (Objects.nonNull(columnAnnotation)) {

                final PxlExportColumnOption columnOption = Optional.ofNullable(columnOptions)
                        .flatMap(options -> options.stream()
                                .filter(o -> StringUtils.equals(o.getFieldName(), columnField.getName()))
                                .findFirst())
                        .orElse(null);

                final List<String> candidateColumnNames = makeCandidateColumnNames(workbookMeta, columnOption, columnAnnotation, columnField);

                // Ignore if the column name is empty.
                if (PxlCollectionUtils.isEmpty(candidateColumnNames)) {
                    continue;
                }

                // Ignore if the column name is already overridden and in use.
                final boolean overriddenColumn = candidateColumnNames.stream()
                        .anyMatch(overriddenColumnNames::contains);
                if (overriddenColumn) {
                    continue;
                }

                final boolean exportEnabled = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportEnabled()))
                        .orElseGet(columnAnnotation::exportEnabled);
                final boolean exportSampleEnabled = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportSampleEnabled()))
                        .orElseGet(columnAnnotation::exportSampleEnabled);

                // Translated below, once the collection separator this sample is split on is resolved.
                String exportSample = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportSample()))
                        .orElseGet(columnAnnotation::exportSample);

                final boolean exportTrim = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportTrim()))
                        .orElseGet(columnAnnotation::exportTrim);
                String exportPattern = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportPattern()))
                        .orElseGet(columnAnnotation::exportPattern);
                if (StringUtils.isBlank(exportPattern)) {
                    exportPattern = columnAnnotation.pattern();
                }
                final int exportColumnWidth = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportColumnWidth()))
                        .orElseGet(columnAnnotation::exportColumnWidth);
                String exportCollectionSeparator = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportCollectionSeparator()))
                        .orElseGet(columnAnnotation::exportCollectionSeparator);
                if (StringUtils.isEmpty(exportCollectionSeparator)) {
                    exportCollectionSeparator = columnAnnotation.collectionSeparator();
                }
                if (StringUtils.isEmpty(exportCollectionSeparator)) {
                    exportCollectionSeparator = PxlConstants.DEFAULT_COLLECTION_SEPARATOR;
                }

                // The sample of a String/enum column is a content-i18n key as well. A Collection of those carries one
                // key per element, so it is split on the separator resolved just above and translated element by element.
                if (exportSampleEnabled
                        && Objects.nonNull(exportResourceBundle)
                        && StringUtils.isNotBlank(exportSample)) {

                    final Class<?> sampleContentClass = resolveContentClass(columnField);
                    if (sampleContentClass == String.class || sampleContentClass.isEnum()) {
                        exportSample = PxlClassSupport.isCollectionClass(columnField.getType())
                                ? translateCollectionSample(exportResourceBundle, exportSample, exportCollectionSeparator)
                                : PxlI18nContent.translate(exportResourceBundle, exportSample);
                    }
                }

                final boolean exportOverrideSuperClassColumn = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportOverrideSuperClassColumn()))
                        .orElseGet(columnAnnotation::exportOverrideSuperClassColumn);
                final String exportOrder = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportOrder()))
                        .orElseGet(columnAnnotation::exportOrder);
                final String exportMasking = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportMasking()))
                        .orElseGet(columnAnnotation::exportMasking);
                String[] exportOptionItems = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportOptionItems()))
                        .orElseGet(columnAnnotation::exportOptionItems);

                // A String column is the only one whose dropdown is translated, because it is the only one that writes
                // translated text into the cell: an enum, numeric or temporal column always writes its canonical value,
                // so translating its items would leave what is written outside the list it is validated against.
                if ((exportEnabled || exportSampleEnabled)
                        && Objects.nonNull(exportResourceBundle)
                        && ArrayUtils.isNotEmpty(exportOptionItems)
                        && resolveContentClass(columnField) == String.class) {

                    exportOptionItems = Arrays.stream(exportOptionItems)
                            .map(optionItem -> PxlI18nContent.translate(exportResourceBundle, optionItem))
                            .toArray(String[]::new);
                }

                final PxlColumn.ExportEnumDropDownListStyle exportEnumDropDownListStyle = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportEnumDropDownListStyle()))
                        .orElseGet(columnAnnotation::exportEnumDropDownListStyle);
                final String exportNullString = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportNullString()))
                        .orElseGet(columnAnnotation::exportNullString);
                final String exportTrueString = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportTrueString()))
                        .orElseGet(columnAnnotation::exportTrueString);
                final String exportFalseString = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportFalseString()))
                        .orElseGet(columnAnnotation::exportFalseString);
                final boolean exportStringAsPicture = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportStringAsPicture()))
                        .orElseGet(columnAnnotation::exportStringAsPicture);
                final boolean exportStringAsFormula = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getExportStringAsFormula()))
                        .orElseGet(columnAnnotation::exportStringAsFormula);

                Class<? extends PxlStyler> exportColumnRequiredHeaderCellStyler =
                        Objects.nonNull(columnOption) && Objects.nonNull(columnOption.getExportColumnRequiredHeaderCellStyler())
                                ? columnOption.getExportColumnRequiredHeaderCellStyler()
                                : columnAnnotation.exportColumnRequiredHeaderCellStyler();
                if (!PxlMiscUtils.isEffectiveCellStylerClass(exportColumnRequiredHeaderCellStyler)) {
                    exportColumnRequiredHeaderCellStyler = sheetMeta.getExportSheetRequiredHeaderCellStyler();
                }

                Class<? extends PxlStyler> exportColumnOptionalHeaderCellStyler =
                        Objects.nonNull(columnOption) && Objects.nonNull(columnOption.getExportColumnOptionalHeaderCellStyler())
                                ? columnOption.getExportColumnOptionalHeaderCellStyler()
                                : columnAnnotation.exportColumnOptionalHeaderCellStyler();
                if (!PxlMiscUtils.isEffectiveCellStylerClass(exportColumnOptionalHeaderCellStyler)) {
                    exportColumnOptionalHeaderCellStyler = sheetMeta.getExportSheetOptionalHeaderCellStyler();
                }

                Class<? extends PxlStyler> exportColumnDataCellStyler =
                        Objects.nonNull(columnOption) && Objects.nonNull(columnOption.getExportColumnDataCellStyler())
                                ? columnOption.getExportColumnDataCellStyler()
                                : columnAnnotation.exportColumnDataCellStyler();
                if (!PxlMiscUtils.isEffectiveCellStylerClass(exportColumnDataCellStyler)) {
                    exportColumnDataCellStyler = sheetMeta.getExportSheetDataCellStyler();
                }

                columnField.setAccessible(true);
                columnMetas.add(
                        new PxlExportColumnMeta(
                                workbookMeta,                       // workbookMeta
                                sheetMeta,                          // sheetMeta
                                columnField,                        // columnField
                                candidateColumnNames,               // candidateColumnNames
                                exportEnabled,                      // exportEnabled
                                exportSampleEnabled,                // exportSampleEnabled
                                exportSample,                       // exportSample
                                exportTrim,                         // exportTrim
                                exportPattern,                      // exportPattern
                                exportColumnWidth,                  // exportColumnWidth
                                exportCollectionSeparator,          // exportCollectionSeparator
                                exportOverrideSuperClassColumn,     // exportOverrideSuperClassColumn
                                exportOrder,                        // exportOrder
                                exportMasking,                      // exportMasking
                                exportOptionItems,                  // exportOptionItems
                                exportEnumDropDownListStyle,        // exportEnumDropDownListStyle
                                exportNullString,                   // exportNullString
                                exportTrueString,                   // exportTrueString
                                exportFalseString,                  // exportFalseString
                                exportStringAsPicture,              // exportStringAsPicture
                                exportStringAsFormula,              // exportStringAsFormula
                                exportColumnRequiredHeaderCellStyler,     // exportColumnRequiredHeaderCellStyler
                                exportColumnOptionalHeaderCellStyler,      // exportColumnOptionalHeaderCellStyler
                                exportColumnDataCellStyler                // exportColumnDataCellStyler
                        ));

                if (exportEnabled && exportOverrideSuperClassColumn) {
                    overriddenColumnNames.addAll(candidateColumnNames);
                }
            }
        }

        columnMetas.sort(Comparator.comparing(PxlExportColumnMeta::getExportOrder));

        final int defaultFirstColumnIndex = 0;
        final int numOfColumns = PxlCollectionUtils.size(columnMetas);
        int actualExportOriginDataColumnIndex = sheetMeta.getExportFirstDataColumnIndex();
        if (actualExportOriginDataColumnIndex == PxlConstants.DEFAULT_EXPORT_FIRST_DATA_COLUMN_INDEX) {
            actualExportOriginDataColumnIndex = defaultFirstColumnIndex;
        } else {
            actualExportOriginDataColumnIndex -= 1;  // specified as 1-based, so convert to 0-based.
            actualExportOriginDataColumnIndex = Math.max(actualExportOriginDataColumnIndex, defaultFirstColumnIndex);
        }

        int actualExportBoundDataColumnIndex = sheetMeta.getExportLastDataColumnIndex();
        if (actualExportBoundDataColumnIndex == PxlConstants.DEFAULT_EXPORT_LAST_DATA_COLUMN_INDEX) {
            actualExportBoundDataColumnIndex = actualExportOriginDataColumnIndex + numOfColumns;
        } else {
            actualExportBoundDataColumnIndex -= 1;  // specified as 1-based, so convert to 0-based.
            actualExportBoundDataColumnIndex += 1;  // add 1 to use it as an exclusive bound.
            actualExportBoundDataColumnIndex = Math.min(actualExportBoundDataColumnIndex, actualExportOriginDataColumnIndex + numOfColumns);
        }

        int exportColumnIndex = actualExportOriginDataColumnIndex;
        for (final PxlExportColumnMeta columnMeta : columnMetas) {
            if (!isForSample && !columnMeta.isExportEnabled()) {
                continue;
            }
            if (isForSample && !columnMeta.isExportSampleEnabled()) {
                continue;
            }

            columnMeta.setActualExportColumnIndex(exportColumnIndex);
            columnMeta.setActualExportColumnName(PxlCollectionUtils.get(columnMeta.getCandidateColumnNames(), 0));  // export using the first column name.

            exportColumnIndex++;
            if (exportColumnIndex >= actualExportBoundDataColumnIndex) {
                break;
            }
        }

        final long numOfExportedColumns = columnMetas.stream().filter(columnMeta -> columnMeta.getActualExportColumnIndex() >= 0).count();
        final int maxNumOfColumns = workbookMeta.getExportFileFormat().getMaxExportColumns();
        if (numOfExportedColumns > maxNumOfColumns) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_COLUMN_COUNT_EXCEEDED, sheetMeta.getActualExportSheetName(), String.valueOf(maxNumOfColumns)));
        }

        // if the column range specified by exportLastDataColumnIndex is smaller than the number of columns to export, some columns cannot be assigned
        final boolean hasDroppedColumn = columnMetas.stream()
                .filter(columnMeta -> isForSample ? columnMeta.isExportSampleEnabled() : columnMeta.isExportEnabled())
                .anyMatch(columnMeta -> columnMeta.getActualExportColumnIndex() < 0);
        if (hasDroppedColumn) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_COLUMN_RANGE_TRUNCATED, sheetMeta.getActualExportSheetName()));
        }

        // check for duplicate export column names
        final List<String> exportedColumnNames = columnMetas.stream()
                .filter(columnMeta -> columnMeta.getActualExportColumnIndex() >= 0)
                .map(PxlExportColumnMeta::getActualExportColumnName)
                .collect(Collectors.toList());
        final Set<String> duplicatedColumnNames = PxlCollectionUtils.findDuplicates(exportedColumnNames);
        if (!duplicatedColumnNames.isEmpty()) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_DUPLICATE_COLUMN_NAME, sheetMeta.getActualExportSheetName(), duplicatedColumnNames));
        }

        sheetMeta.setActualExportOriginDataColumnIndex(actualExportOriginDataColumnIndex);
        sheetMeta.setActualExportBoundDataColumnIndex(exportColumnIndex);

        return columnMetas;
    }

    /**
     * Returns whether this column's value is rendered to a string cell, i.e. an export pattern or masking pattern applies.
     *
     * @return {@code true} if an export pattern or masking pattern is set
     */
    public boolean isExportedToString() {

        return StringUtils.isNotBlank(exportPattern) || Objects.nonNull(exportMaskingPattern);
    }

    /**
     * Resolves the type a column actually writes: the element type of a Collection column, the field type otherwise.
     * (export)
     * <p>
     * A {@code List<String>} field is a {@code List} by declaration but writes strings, so the element type is what
     * decides whether a value of this column can be translated. Mirrors how the codecs pick a Collection column's
     * element type.
     *
     * @param columnField the column field to resolve
     * @return the element type for a Collection column, the field type otherwise
     * @throws PxlReflectionException if the column is a raw Collection or its element type is not a concrete class
     */
    private static Class<?> resolveContentClass(final Field columnField)
            throws PxlReflectionException {

        return PxlClassSupport.isCollectionClass(columnField.getType())
                ? PxlReflectionSupport.getParameterizedArgument0(columnField)
                : columnField.getType();
    }

    /**
     * Translates a Collection column's {@code exportSample} one element at a time. (export)
     * <p>
     * The sample holds the elements joined by the column's export collection separator, the same form the collection
     * codec splits back apart, so each element is translated on its own and the results are joined with that separator
     * again. Translating the sample as a single key would instead put the separator inside the bundle value, where a
     * change to {@code exportCollectionSeparator} could no longer reach it. Elements missing from the bundle pass
     * through unchanged, and empty elements are preserved so the element count is kept.
     *
     * @param exportResourceBundle      the consumer bundle to translate through; {@code null} leaves the sample unchanged
     * @param exportSample              the sample value holding the elements joined by the separator
     * @param exportCollectionSeparator the separator the elements are joined with
     * @return the sample with each of its elements translated
     */
    private static String translateCollectionSample(@Nullable final ResourceBundle exportResourceBundle,
                                                    final String exportSample,
                                                    final String exportCollectionSeparator) {

        final String[] sampleElements = StringUtils.splitByWholeSeparatorPreserveAllTokens(exportSample, exportCollectionSeparator);

        return Arrays.stream(sampleElements)
                .map(sampleElement -> PxlI18nContent.translate(exportResourceBundle, sampleElement))
                .collect(Collectors.joining(exportCollectionSeparator));
    }

    /**
     * Resolves the ordered candidate column names, preferring the column option's names, then the {@link PxlColumn}
     * annotation names, and finally the field name; each name is i18n-translated and blank entries are dropped.
     *
     * @param workbookMeta     the workbook metadata supplying the export resource bundle
     * @param columnOption     the column option whose names take priority; may be {@code null}
     * @param columnAnnotation the column annotation providing fallback names
     * @param columnField      the column field whose name is the final fallback
     * @return the non-empty ordered list of candidate column names
     * @throws PxlNullPointerException if {@code workbookMeta}, {@code columnAnnotation}, or {@code columnField} is {@code null}
     */
    private static List<String> makeCandidateColumnNames(final PxlExportWorkbookMeta workbookMeta,
                                                         @Nullable final PxlExportColumnOption columnOption,
                                                         final PxlColumn columnAnnotation,
                                                         final Field columnField)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");
        PxlAssertSupport.notNull(columnAnnotation, "columnAnnotation");
        PxlAssertSupport.notNull(columnField, "columnField");

        final ResourceBundle exportResourceBundle = Optional.ofNullable(workbookMeta)
                .map(PxlExportWorkbookMeta::getExportResourceBundle)
                .orElse(null);

        List<String> candidateColumnNames;

        candidateColumnNames = PxlExportColumnOption.getExportColumnNames(columnOption);
        if (PxlCollectionUtils.isNotEmpty(candidateColumnNames)) {
            candidateColumnNames = candidateColumnNames.stream()
                    .map(name -> PxlI18nContent.translate(exportResourceBundle, name))
                    // .map(StringUtils::trim)
                    .filter(StringUtils::isNotBlank)
                    .collect(Collectors.toList());
            if (PxlCollectionUtils.isNotEmpty(candidateColumnNames)) {
                return candidateColumnNames;
            }
        }

        // Get the value of the @PxlColumn annotation, i.e. the column name.
        candidateColumnNames = Arrays.stream(columnAnnotation.name())
                .map(name -> PxlI18nContent.translate(exportResourceBundle, name))
                .map(StringUtils::trim)
                .filter(StringUtils::isNotBlank)
                .collect(Collectors.toList());
        if (PxlCollectionUtils.isNotEmpty(candidateColumnNames)) {
            return candidateColumnNames;
        }

        candidateColumnNames = Arrays.asList(columnField.getName());

        return candidateColumnNames;
    }

    /**
     * Writes the value to the cell as quote-prefixed text, delegating to the workbook metadata.
     *
     * @param cell  the target cell
     * @param value the string value to write
     * @return the same cell
     */
    public Cell setQuotePrefixedCellValue(final Cell cell,
                                          final String value) {

        return this.workbookMeta.setQuotePrefixedCellValue(cell, value);
    }

    /**
     * Resolved custom export converter for a column value type: an optional {@link PxlExportConverter}-annotated
     * method and the type's {@code toString} method, used to render the value to a String on export.
     */
    @Getter
    public static final class PxlExportConverterMeta {

        private final Class<?> valueClass;

        private final Method exportConverterMethod;

        private final Method toStringMethod;

        /**
         * Creates the export converter metadata holding the resolved converter members for a value type.
         *
         * @param valueClass            the value type
         * @param exportConverterMethod the {@link PxlExportConverter}-annotated method, or {@code null}
         * @param toStringMethod        the type's {@code toString} method, or {@code null}
         */
        private PxlExportConverterMeta(final Class<?> valueClass,
                                       final Method exportConverterMethod,
                                       final Method toStringMethod) {

            this.valueClass = valueClass;
            this.exportConverterMethod = exportConverterMethod;
            this.toStringMethod = toStringMethod;
        }

        /**
         * Resolves the export converter metadata for the given value type, validating any {@link PxlExportConverter} method.
         *
         * @param objectClass the value type to resolve a converter for
         * @return the resolved export converter metadata
         * @throws PxlArgumentException if the converter method has an invalid signature/return type, or the type has neither a converter nor a {@code toString}
         */
        public static PxlExportConverterMeta of(final Class<?> objectClass)
                throws PxlArgumentException {

            final Method exportConverterMethod = PxlReflectionSupport.getAnnotatedMethod(objectClass, PxlExportConverter.class);
            final Method toStringMethod = PxlReflectionSupport.getToStringMethod(objectClass);

            if (Objects.nonNull(exportConverterMethod)) {
                if (exportConverterMethod.getReturnType() != String.class) {
                    throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_EXPORT_CONVERTER_RETURN_STRING, objectClass.getSimpleName()));
                }
                if (Modifier.isStatic(exportConverterMethod.getModifiers())) {
                    if (exportConverterMethod.getParameterCount() != 1
                            || !exportConverterMethod.getParameterTypes()[0].isAssignableFrom(objectClass)) {
                        throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_EXPORT_CONVERTER_STATIC_ONE_PARAM, objectClass.getSimpleName()));
                    }
                } else {
                    if (exportConverterMethod.getParameterCount() != 0) {
                        throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_EXPORT_CONVERTER_INSTANCE_NO_PARAM, objectClass.getSimpleName()));
                    }
                }
            }

            if (ObjectUtils.allNull(exportConverterMethod, toStringMethod)) {
                throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_COLUMN_TYPE_UNSUPPORTED, objectClass.getSimpleName()));
            }

            return new PxlExportConverterMeta(objectClass, exportConverterMethod, toStringMethod);
        }

    }

}
