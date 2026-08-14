package io.github.hclimkr.pxl.internal.meta;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.annotation.PxlImportConverter;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.exception.PxlReflectionException;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nContent;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.support.*;
import io.github.hclimkr.pxl.option.PxlImportColumnOption;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import lombok.Getter;
import lombok.Setter;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;

import javax.validation.constraints.NotBlank;
import javax.validation.constraints.NotEmpty;
import javax.validation.constraints.NotNull;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Column import metadata, resolved for an Excel or a CSV source alike.
 */
@Getter
public final class PxlImportColumnMeta {

    private final PxlImportWorkbookMeta workbookMeta;

    private final PxlImportSheetMeta sheetMeta;

    private final Field columnField;

    private final Class<?> columnClass;

    private final List<String> candidateColumnNames;

    private final boolean importEnabled;

    private final boolean importTrim;

    private final boolean importUnique;

    private final String importPattern;

    private final String importTrueString;

    private final String importFalseString;

    private final String importCollectionSeparator;

    private final boolean importOverrideSuperClassColumn;

    private final boolean isRequired;

    private DecimalFormat importDecimalFormatterCache = null;

    private SimpleDateFormat importJavaDateFormatterCache = null;

    private DateTimeFormatter importDateTimeFormatterCache = null;

    private PxlTemporalAmountSupport.CompiledTemporalPattern importTemporalPatternCache = null;

    private PxlImportConverterMeta importCustomConverterMeta = null;

    @Setter
    private String actualImportColumnName;

    @Setter
    private int actualImportColumnIndex = -1;

    /**
     * Creates the resolved import metadata for one column, storing the merged option/annotation values and,
     * when import is enabled, pre-compiling the import number/date/temporal pattern and custom converter.
     *
     * @param workbookMeta                   the enclosing workbook metadata
     * @param sheetMeta                      the enclosing sheet metadata
     * @param columnField                    the column field to bind
     * @param candidateColumnNames           the ordered candidate column names to match against the header
     * @param importEnabled                  whether this column is imported
     * @param importTrim                     whether string values are trimmed on import
     * @param importUnique                   whether this column's values must be unique
     * @param importPattern                  the number/date/temporal parse pattern; blank for none
     * @param importTrueString               the token parsed as boolean {@code true} (kept only for String/Boolean columns)
     * @param importFalseString              the token parsed as boolean {@code false} (kept only for String/Boolean columns)
     * @param importCollectionSeparator      the element separator used when parsing Collection columns
     * @param importOverrideSuperClassColumn whether this column overrides a same-named super-class column
     * @throws PxlReflectionException if the Collection element type cannot be resolved
     * @throws PxlArgumentException   if the import pattern is malformed or a custom converter has an invalid signature
     */
    private PxlImportColumnMeta(final PxlImportWorkbookMeta workbookMeta,
                                final PxlImportSheetMeta sheetMeta,
                                final Field columnField,
                                final List<String> candidateColumnNames,
                                final boolean importEnabled,
                                final boolean importTrim,
                                final boolean importUnique,
                                final String importPattern,
                                final String importTrueString,
                                final String importFalseString,
                                final String importCollectionSeparator,
                                final boolean importOverrideSuperClassColumn)
            throws PxlReflectionException, PxlArgumentException {

        this.workbookMeta = workbookMeta;
        this.sheetMeta = sheetMeta;

        this.columnField = columnField;
        this.columnClass = columnField.getType();

        this.candidateColumnNames = candidateColumnNames;
        this.importEnabled = importEnabled;
        this.importTrim = importTrim;
        this.importUnique = importUnique;
        this.importPattern = importPattern;
        this.importTrueString = (PxlClassSupport.isStringClass(this.columnClass) || PxlClassSupport.isBooleanClass(this.columnClass)) ? importTrueString : null;
        this.importFalseString = (PxlClassSupport.isStringClass(this.columnClass) || PxlClassSupport.isBooleanClass(this.columnClass)) ? importFalseString : null;
        this.importCollectionSeparator = importCollectionSeparator;
        this.importOverrideSuperClassColumn = importOverrideSuperClassColumn;

        this.isRequired = (Objects.nonNull(columnField.getAnnotation(NotNull.class)))
                || (Objects.nonNull(columnField.getAnnotation(NotEmpty.class)))
                || (Objects.nonNull(columnField.getAnnotation(NotBlank.class)));

        if (importEnabled) {
            if (StringUtils.isNotBlank(importPattern)) {
                // For a Collection column, the element type is the target of the pattern.
                final Class<?> patternTargetClass = PxlClassSupport.isCollectionClass(this.columnClass)
                        ? PxlReflectionSupport.getParameterizedArgument0(this.columnField)
                        : this.columnClass;

                if (PxlClassSupport.isNumberClass(patternTargetClass)) {
                    try {
                        this.importDecimalFormatterCache = PxlNumberSupport.getDecimalFormat(importPattern);
                    } catch (IllegalArgumentException illegalArgumentException) {
                        throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_IMPORT_NUMBER_PATTERN_INVALID, importPattern), illegalArgumentException);
                    }

                    if (patternTargetClass == BigDecimal.class || patternTargetClass == BigInteger.class) {
                        this.importDecimalFormatterCache.setParseBigDecimal(true);
                    }
                } else if (PxlClassSupport.isJavaDateClass(patternTargetClass)) {
                    try {
                        this.importJavaDateFormatterCache = PxlDateTimeSupport.getCellSimpleDateFormatter(importPattern, Locale.ROOT);
                    } catch (IllegalArgumentException illegalArgumentException) {
                        throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_IMPORT_DATE_TIME_PATTERN_INVALID, importPattern), illegalArgumentException);
                    }
                } else if (PxlClassSupport.isDateTimeClass(patternTargetClass)) {
                    try {
                        this.importDateTimeFormatterCache = PxlDateTimeSupport.getCellDateTimeFormatter(importPattern, Locale.ROOT);
                    } catch (IllegalArgumentException illegalArgumentException) {
                        throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_IMPORT_DATE_TIME_PATTERN_INVALID, importPattern), illegalArgumentException);
                    }
                } else if (PxlClassSupport.isTemporalAmountClass(patternTargetClass)) {
                    this.importTemporalPatternCache = PxlTemporalAmountSupport.compileTemporalPattern(importPattern);
                }
            }

            final Class<?> targetClass;
            if (PxlClassSupport.isCollectionClass(this.columnClass)) {
                targetClass = PxlReflectionSupport.getParameterizedArgument0(this.columnField);
            } else {
                targetClass = this.columnClass;
            }

            if (PxlClassSupport.isCustomConvertableClass(targetClass)) {
                this.importCustomConverterMeta = PxlImportConverterMeta.of(targetClass);
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
     * Returns whether this column uses a custom import converter ({@link PxlImportConverter}, a String constructor, or an enum).
     *
     * @return {@code true} if a custom import converter is resolved for this column
     */
    public boolean isImportCustomConvertable() {

        return Objects.nonNull(importCustomConverterMeta);
    }

    /**
     * On import, collects the sheet's column metadata from the row class.
     * Each {@link PxlColumn}-annotated field becomes one column; column options (matched by field name) override the
     * annotation and i18n-resolved candidate names are computed.
     *
     * @param sheetMeta the enclosing sheet metadata, supplying the row class and column options
     * @return the collected column metadata list
     * @throws PxlNullPointerException if {@code sheetMeta} is {@code null}
     * @throws PxlReflectionException  if a column field's type cannot be resolved
     * @throws PxlArgumentException    if an import pattern is malformed or a custom converter has an invalid signature
     */
    public static List<PxlImportColumnMeta> makeImportColumnMetas(final PxlImportSheetMeta sheetMeta)
            throws PxlNullPointerException, PxlReflectionException, PxlArgumentException {

        PxlAssertSupport.notNull(sheetMeta, "sheetMeta");

        final Class<?> rowClass = sheetMeta.getRowClass();
        final PxlImportWorkbookMeta workbookMeta = sheetMeta.getWorkbookMeta();
        final List<PxlImportColumnOption> columnOptions = sheetMeta.getImportColumnOptions();

//        final Field[] columnFields = rowClass.getDeclaredFields();
        final List<Field> columnFields = PxlReflectionSupport.getAllFields(rowClass);
        final List<PxlImportColumnMeta> columnMetas = new ArrayList<>(PxlCollectionUtils.size(columnFields));
        final Set<String> overriddenColumnNames = new HashSet<>();

        for (final Field columnField : columnFields) {
            final PxlColumn columnAnnotation = columnField.getAnnotation(PxlColumn.class);

            if (Objects.nonNull(columnAnnotation)) {

                final PxlImportColumnOption columnOption = Optional.ofNullable(columnOptions)
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

                final boolean importEnabled = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getImportEnabled()))
                        .orElseGet(columnAnnotation::importEnabled);
                final boolean importTrim = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getImportTrim()))
                        .orElseGet(columnAnnotation::importTrim);
                final boolean importUnique = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getImportUnique()))
                        .orElseGet(columnAnnotation::importUnique);
                String importPattern = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getImportPattern()))
                        .orElseGet(columnAnnotation::importPattern);
                if (StringUtils.isBlank(importPattern)) {
                    importPattern = columnAnnotation.pattern();
                }
                final String importTrueString = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getImportTrueString()))
                        .orElseGet(columnAnnotation::importTrueString);
                final String importFalseString = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getImportFalseString()))
                        .orElseGet(columnAnnotation::importFalseString);
                String importCollectionSeparator = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getImportCollectionSeparator()))
                        .orElseGet(columnAnnotation::importCollectionSeparator);
                if (StringUtils.isEmpty(importCollectionSeparator)) {
                    importCollectionSeparator = columnAnnotation.collectionSeparator();
                }
                if (StringUtils.isEmpty(importCollectionSeparator)) {
                    importCollectionSeparator = PxlConstants.DEFAULT_COLLECTION_SEPARATOR;
                }
                final boolean importOverrideSuperClassColumn = Optional.ofNullable(columnOption)
                        .flatMap(option -> Optional.ofNullable(option.getImportOverrideSuperClassColumn()))
                        .orElseGet(columnAnnotation::importOverrideSuperClassColumn);

                columnField.setAccessible(true);
                columnMetas.add(
                        new PxlImportColumnMeta(
                                workbookMeta,                   // workbookMeta
                                sheetMeta,                      // sheetMeta
                                columnField,                    // columnField
                                candidateColumnNames,           // candidateColumnNames
                                importEnabled,                  // importEnabled
                                importTrim,                     // importTrim
                                importUnique,                   // importUnique
                                importPattern,                  // importPattern
                                importTrueString,               // importTrueString
                                importFalseString,              // importFalseString
                                importCollectionSeparator,      // importCollectionSeparator
                                importOverrideSuperClassColumn  // importOverrideSuperClassColumn
                        ));

                if (importEnabled && importOverrideSuperClassColumn) {
                    overriddenColumnNames.addAll(candidateColumnNames);
                }
            }
        }

        return columnMetas;
    }

    /**
     * Resolves the ordered candidate column names, preferring the column option's names, then the {@link PxlColumn}
     * annotation names, and finally the field name; each name is i18n-translated, whitespace-stripped and blank entries dropped.
     *
     * @param workbookMeta     the workbook metadata supplying the import resource bundle
     * @param columnOption     the column option whose names take priority; may be {@code null}
     * @param columnAnnotation the column annotation providing fallback names
     * @param columnField      the column field whose name is the final fallback
     * @return the non-empty ordered list of candidate column names
     * @throws PxlNullPointerException if {@code workbookMeta}, {@code columnAnnotation}, or {@code columnField} is {@code null}
     */
    private static List<String> makeCandidateColumnNames(final PxlImportWorkbookMeta workbookMeta,
                                                         @Nullable final PxlImportColumnOption columnOption,
                                                         final PxlColumn columnAnnotation,
                                                         final Field columnField)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(workbookMeta, "workbookMeta");
        PxlAssertSupport.notNull(columnAnnotation, "columnAnnotation");
        PxlAssertSupport.notNull(columnField, "columnField");

        final ResourceBundle importResourceBundle = Optional.ofNullable(workbookMeta)
                .map(PxlImportWorkbookMeta::getImportResourceBundle)
                .orElse(null);

        List<String> candidateColumnNames;

        candidateColumnNames = PxlImportColumnOption.getImportColumnNames(columnOption);
        if (PxlCollectionUtils.isNotEmpty(candidateColumnNames)) {
            candidateColumnNames = candidateColumnNames.stream()
                    .map(name -> PxlI18nContent.translate(importResourceBundle, name))
                    .map(StringUtils::deleteWhitespace)
                    .filter(StringUtils::isNotBlank)
                    .collect(Collectors.toList());
            if (PxlCollectionUtils.isNotEmpty(candidateColumnNames)) {
                return candidateColumnNames;
            }
        }

        // Get the value of the @PxlColumn annotation, i.e. the column name.
        candidateColumnNames = Arrays.stream(columnAnnotation.name())
                .map(name -> PxlI18nContent.translate(importResourceBundle, name))
                .map(StringUtils::deleteWhitespace)
                .filter(StringUtils::isNotBlank)
                .collect(Collectors.toList());
        if (PxlCollectionUtils.isNotEmpty(candidateColumnNames)) {
            return candidateColumnNames;
        }

        candidateColumnNames = Arrays.asList(columnField.getName());

        return candidateColumnNames;
    }

    /**
     * Resolved custom import converter for a column value type: an optional {@link PxlImportConverter}-annotated static
     * method, a String constructor and the type's {@code toString} method, used to parse a String into the value on import.
     */
    @Getter
    public static final class PxlImportConverterMeta {

        private final Class<?> valueClass;

        private final Method importConverterMethod;

        private final Constructor<?> stringConstructor;

        private final Method toStringMethod;

        /**
         * Creates the import converter metadata holding the resolved converter members for a value type.
         *
         * @param valueClass            the value type
         * @param importConverterMethod the {@link PxlImportConverter}-annotated method, or {@code null}
         * @param stringConstructor     the single-String constructor, or {@code null}
         * @param toStringMethod        the type's {@code toString} method, or {@code null}
         */
        private PxlImportConverterMeta(final Class<?> valueClass,
                                       final Method importConverterMethod,
                                       final Constructor<?> stringConstructor,
                                       final Method toStringMethod) {

            this.valueClass = valueClass;
            this.importConverterMethod = importConverterMethod;
            this.stringConstructor = stringConstructor;
            this.toStringMethod = toStringMethod;
        }

        /**
         * Resolves the import converter metadata for the given value type, validating any {@link PxlImportConverter} method.
         *
         * @param objectClass the value type to resolve a converter for
         * @return the resolved import converter metadata
         * @throws PxlArgumentException if the converter method has an invalid signature/return type, or a non-enum type has neither a converter nor a String constructor
         */
        public static PxlImportConverterMeta of(final Class<?> objectClass)
                throws PxlArgumentException {

            final Method importConverterMethod = PxlReflectionSupport.getAnnotatedMethod(objectClass, PxlImportConverter.class);
            final Constructor<?> stringConstructor = PxlReflectionSupport.getStringTypeConstructor(objectClass);
            final Method toStringMethod = PxlReflectionSupport.getToStringMethod(objectClass);

            if (Objects.nonNull(importConverterMethod)) {
                if (importConverterMethod.getReturnType() != objectClass) {
                    throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_IMPORT_CONVERTER_RETURN_TYPE, objectClass.getSimpleName()));
                }
                if (!Modifier.isStatic(importConverterMethod.getModifiers())) {
                    throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_IMPORT_CONVERTER_STATIC, objectClass.getSimpleName()));
                }
                if (importConverterMethod.getParameterCount() != 1
                        || !importConverterMethod.getParameterTypes()[0].isAssignableFrom(String.class)) {
                    throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_IMPORT_CONVERTER_ONE_STRING_PARAM, objectClass.getSimpleName()));
                }
            }

            if (!objectClass.isEnum() && ObjectUtils.allNull(importConverterMethod, stringConstructor)) {
                throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.META_COLUMN_TYPE_UNSUPPORTED, objectClass.getSimpleName()));
            }

            return new PxlImportConverterMeta(objectClass, importConverterMethod, stringConstructor, toStringMethod);
        }

    }

}
