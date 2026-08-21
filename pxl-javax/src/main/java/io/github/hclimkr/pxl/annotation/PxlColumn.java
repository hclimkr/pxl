package io.github.hclimkr.pxl.annotation;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.styler.PxlStyler;

import java.lang.annotation.*;

/**
 * Annotation for specifying properties on an Excel column
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface PxlColumn {

    /**
     * Specifies the Excel column name.
     * <p>
     * The value doubles as a content-i18n key: when the workbook sets {@code exportI18nBaseName} /
     * {@code importI18nBaseName}, the name is resolved through that bundle first, and it is the translation that is
     * written as the header on export and matched against it on import. A name the bundle does not carry is used as
     * it stands.
     *
     * @return the column name(s); when empty ({@code {}}), the field name is used
     */
    String[] name() default {};

    /**
     * Specifies the cell formatting string used when exporting/importing the Excel column.
     * Each direction falls back to this value when its own pattern - {@link #exportPattern()},
     * {@link #importPattern()} - is empty.
     * <p>
     * A pattern has to match the value it parses in full. A value the pattern reads only the front of - {@code "1e3"}
     * or {@code "123abc"} under {@code "#,##0"}, {@code "2024-01-02 xxx"} under {@code "yyyy-MM-dd"} - is rejected
     * rather than bound as the part that could be read. Naming a pattern therefore never makes a column accept more
     * than it would without one.
     *
     * @return the shared export/import cell-formatting pattern; empty ({@code ""}) by default
     */
    String pattern() default "";

    /**
     * Specifies the separator between each Element when exporting/importing a Collection.
     * Each direction falls back to this value when its own separator - {@link #exportCollectionSeparator()},
     * {@link #importCollectionSeparator()} - is empty.
     *
     * @return the shared collection element separator; defaults to {@link PxlConstants#DEFAULT_COLLECTION_SEPARATOR} ({@code ";"})
     */
    String collectionSeparator() default PxlConstants.DEFAULT_COLLECTION_SEPARATOR;

    /**
     * Specifies whether to export.
     *
     * @return {@code true} to write this column on export; {@code true} by default
     */
    boolean exportEnabled() default true;

    /**
     * Specifies whether to export a sample.
     *
     * @return {@code true} to include this column in a sample export; {@code true} by default
     */
    boolean exportSampleEnabled() default true;

    /**
     * Specifies the sample value on export.
     * <p>
     * On a String or enum column the value is also a content-i18n key, so a workbook that declares
     * {@code exportI18nBaseName} writes the translation instead. A Collection of those types holds one key per
     * element, separated by {@link #exportCollectionSeparator()}, and each element is translated on its own.
     * An enum sample is parsed back into its constant before being written, so the cell ends up holding the
     * canonical value rather than the translated text.
     *
     * @return the sample cell value written for this column in a sample export; empty ({@code ""}) by default
     */
    String exportSample() default "";

    /**
     * Specifies whether to trim the string on export.
     *
     * @return {@code true} to trim the exported string; defaults to {@link PxlConstants#DEFAULT_EXPORT_TRIM} ({@code false})
     */
    boolean exportTrim() default PxlConstants.DEFAULT_EXPORT_TRIM;

    /**
     * Specifies the cell formatting string on export.
     * Valid only for fields of type Numeric, Date, LocalTime, LocalDate, LocalDateTime, ZonedDateTime, OffsetTime, OffsetDateTime, Duration, Period.
     *
     * @return the export-specific cell-formatting pattern; empty ({@code ""}) falls back to {@link #pattern()}
     */
    String exportPattern() default "";

    /**
     * Specifies the column width on export.
     * in units of 1/256th of a character width (maximum: 255 * 256)
     *
     * @return the export column width in units of 1/256th of a character width; defaults to {@link PxlConstants#DEFAULT_EXPORT_COLUMN_WIDTH} (auto-size)
     */
    int exportColumnWidth() default PxlConstants.DEFAULT_EXPORT_COLUMN_WIDTH;

    /**
     * Specifies the separator between each Element when exporting a Collection.
     *
     * @return the export-specific collection element separator; empty ({@code ""} by default) falls back to {@link #collectionSeparator()}
     */
    String exportCollectionSeparator() default PxlConstants.DEFAULT_EXPORT_COLLECTION_SEPARATOR;

    /**
     * Specifies whether, on export, to override a superclass field that uses the same column name, if one exists.
     *
     * @return {@code true} to override a same-named superclass column on export; defaults to {@link PxlConstants#DEFAULT_EXPORT_OVERRIDE_SUPER_CLASS_COLUMN} ({@code false})
     */
    boolean exportOverrideSuperClassColumn() default PxlConstants.DEFAULT_EXPORT_OVERRIDE_SUPER_CLASS_COLUMN;

    /**
     * Specifies the ordering between columns on export. (in alphabetical order)
     *
     * @return the alphabetical sort key that orders columns on export; empty ({@code ""}) by default
     */
    String exportOrder() default "";

    /**
     * Specifies the masking rule as a Regular Expression on export.
     *
     * @return the regular expression describing the masking rule applied on export; empty ({@code ""}, no masking) by default
     */
    String exportMasking() default "";

    /**
     * Sets the list of selectable options on export.
     * <p>
     * On a String column each item is also a content-i18n key, so the dropdown offers the same text the cells hold.
     * A column of any other type writes its value in canonical form, so its items are used verbatim - translating
     * them would leave the written value outside the list it is validated against.
     *
     * @return the selectable dropdown option items offered on export; empty ({@code {}}) by default
     */
    String[] exportOptionItems() default {};

    /**
     * Sets an Enum field as a DropDownList on export.
     *
     * @return how an Enum field's values are rendered as a dropdown list on export; defaults to {@link ExportEnumDropDownListStyle#SET}
     */
    ExportEnumDropDownListStyle exportEnumDropDownListStyle() default ExportEnumDropDownListStyle.SET;

    /**
     * Specifies the string representing null on export.
     *
     * @return the string written for a {@code null} value on export; defaults to {@link PxlConstants#DEFAULT_EXPORT_NULL_STRING} ({@code ""})
     */
    String exportNullString() default PxlConstants.DEFAULT_EXPORT_NULL_STRING;

    /**
     * Specifies the string representing the boolean value true on export.
     *
     * @return the string written for boolean {@code true} on export; defaults to {@link PxlConstants#DEFAULT_EXPORT_TRUE_STRING}
     */
    String exportTrueString() default PxlConstants.DEFAULT_EXPORT_TRUE_STRING;

    /**
     * Specifies the string representing the boolean value false on export.
     *
     * @return the string written for boolean {@code false} on export; defaults to {@link PxlConstants#DEFAULT_EXPORT_FALSE_STRING}
     */
    String exportFalseString() default PxlConstants.DEFAULT_EXPORT_FALSE_STRING;

    /**
     * On export, interprets the string as a path to an image and applies that image itself to the cell.
     *
     * @return {@code true} to treat the string as an image path and embed that image in the cell on export; defaults to {@link PxlConstants#DEFAULT_EXPORT_STRING_AS_PICTURE} ({@code false})
     */
    boolean exportStringAsPicture() default PxlConstants.DEFAULT_EXPORT_STRING_AS_PICTURE;

    /**
     * On export, interprets the string as a formula and applies the computed result itself to the cell.
     *
     * @return {@code true} to treat the string as a formula and apply its computed result to the cell on export; defaults to {@link PxlConstants#DEFAULT_EXPORT_STRING_AS_FORMULA} ({@code false})
     */
    boolean exportStringAsFormula() default PxlConstants.DEFAULT_EXPORT_STRING_AS_FORMULA;

    /**
     * Specifies the style to apply to a required header cell on export.
     *
     * @return the styler for this column's required header cell on export; {@link PxlStyler} (unset) by default, deferring to the sheet/workbook/built-in styler
     */
    Class<? extends PxlStyler> exportColumnRequiredHeaderCellStyler() default PxlStyler.class;

    /**
     * Specifies the style to apply to an optional header cell on export.
     *
     * @return the styler for this column's optional header cell on export; {@link PxlStyler} (unset) by default, deferring to the sheet/workbook/built-in styler
     */
    Class<? extends PxlStyler> exportColumnOptionalHeaderCellStyler() default PxlStyler.class;

    /**
     * Specifies the style to apply to a Data Cell on export.
     *
     * @return the styler for this column's data cells on export; {@link PxlStyler} (unset) by default, deferring to the sheet/workbook/built-in styler
     */
    Class<? extends PxlStyler> exportColumnDataCellStyler() default PxlStyler.class;

    /**
     * Specifies whether to import.
     *
     * @return {@code true} to bind this column on import; {@code true} by default
     */
    boolean importEnabled() default true;

    /**
     * Specifies whether to trim the string on import.
     *
     * @return {@code true} to trim the imported string; defaults to {@link PxlConstants#DEFAULT_IMPORT_TRIM} ({@code true})
     */
    boolean importTrim() default PxlConstants.DEFAULT_IMPORT_TRIM;

    /**
     * Specifies whether to check the uniqueness of the column values on import.
     *
     * @return {@code true} to enforce uniqueness of the imported column values; defaults to {@link PxlConstants#DEFAULT_IMPORT_UNIQUE} ({@code false})
     */
    boolean importUnique() default PxlConstants.DEFAULT_IMPORT_UNIQUE;

    /**
     * Specifies the cell formatting string on import.
     * Valid only for fields of type Numeric, Date, LocalTime, LocalDate, LocalDateTime, ZonedDateTime, OffsetTime, OffsetDateTime, Duration, Period.
     * A Duration/Period pattern is the {@code DurationFormatUtils} style used on export, and a value that does not match it falls back to ISO-8601.
     * The pattern has to consume the cell value in full - see {@link #pattern()}.
     *
     * @return the import-specific cell-formatting pattern; empty ({@code ""}) falls back to {@link #pattern()}
     */
    String importPattern() default "";

    /**
     * Specifies the string representing the boolean value true on import.
     * A String column renders a BOOLEAN cell as this string, and a Boolean column interprets this string (case-insensitive) as true.
     *
     * @return the string mapped to boolean {@code true} on import; defaults to {@link PxlConstants#DEFAULT_IMPORT_TRUE_STRING}
     */
    String importTrueString() default PxlConstants.DEFAULT_IMPORT_TRUE_STRING;

    /**
     * Specifies the string representing the boolean value false on import.
     * A String column renders a BOOLEAN cell as this string, and a Boolean column interprets this string (case-insensitive) as false.
     *
     * @return the string mapped to boolean {@code false} on import; defaults to {@link PxlConstants#DEFAULT_IMPORT_FALSE_STRING}
     */
    String importFalseString() default PxlConstants.DEFAULT_IMPORT_FALSE_STRING;

    /**
     * Specifies the separator between each Element when importing as a Collection.
     *
     * @return the import-specific collection element separator; empty ({@code ""} by default) falls back to {@link #collectionSeparator()}
     */
    String importCollectionSeparator() default PxlConstants.DEFAULT_IMPORT_COLLECTION_SEPARATOR;

    /**
     * Specifies whether, on import, to override a superclass field that uses the same column name, if one exists.
     *
     * @return {@code true} to override a same-named superclass column on import; defaults to {@link PxlConstants#DEFAULT_IMPORT_OVERRIDE_SUPER_CLASS_COLUMN} ({@code false})
     */
    boolean importOverrideSuperClassColumn() default PxlConstants.DEFAULT_IMPORT_OVERRIDE_SUPER_CLASS_COLUMN;

    /**
     * Rendering style for an Enum field's dropdown list on export.
     */
    enum ExportEnumDropDownListStyle {
        /**
         * Render a dropdown listing the enum values in their declared order.
         */
        SET,
        /**
         * Render a dropdown listing the enum values sorted alphabetically.
         */
        SORTED_SET,
        /**
         * Do not render a dropdown for the enum field.
         */
        NONE,
        ;
    }

}
