package io.github.hclimkr.pxl.annotation;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.styler.PxlStyler;

import java.lang.annotation.*;

/**
 * Annotation for specifying properties on an Excel sheet
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface PxlSheet {

    /**
     * Specifies the name for the Excel sheet.
     * On import, it is matched against the actual sheet name (the CSV file name, extension removed, for a CSV source)
     * after whitespace removal, ignoring case.
     * <p>
     * The value doubles as a content-i18n key: when the workbook sets {@code importI18nBaseName} /
     * {@code exportI18nBaseName}, the name is resolved through that bundle first, and it is the translation that is
     * matched on import and written on export. A name the bundle does not carry is used as it stands.
     *
     * @return the sheet name(s); when empty ({@code {}}), the field name is used
     */
    String[] name() default {};

    /**
     * Specifies whether to import.
     *
     * @return {@code true} to bind this sheet on import; {@code true} by default
     */
    boolean importEnabled() default true;

    /**
     * Specifies whether, on import, to override a superclass field that uses the same sheet name, if one exists.
     * Names differing only in case count as the same sheet name here, as they do when matching a sheet.
     *
     * @return {@code true} to override a same-named superclass sheet on import; defaults to {@link PxlConstants#DEFAULT_IMPORT_OVERRIDE_SUPER_CLASS_SHEET} ({@code false})
     */
    boolean importOverrideSuperClassSheet() default PxlConstants.DEFAULT_IMPORT_OVERRIDE_SUPER_CLASS_SHEET;

    /**
     * Specifies whether to exclude (skip) hidden rows on import.
     *
     * @return {@code true} to skip hidden rows on import; defaults to {@link PxlConstants#DEFAULT_IMPORT_EXCLUDE_HIDDEN_ROWS} ({@code false})
     */
    boolean importExcludeHiddenRows() default PxlConstants.DEFAULT_IMPORT_EXCLUDE_HIDDEN_ROWS;

    /**
     * Specifies whether to exclude (skip) hidden columns on import.
     *
     * @return {@code true} to skip hidden columns on import; defaults to {@link PxlConstants#DEFAULT_IMPORT_EXCLUDE_HIDDEN_COLUMNS} ({@code false})
     */
    boolean importExcludeHiddenColumns() default PxlConstants.DEFAULT_IMPORT_EXCLUDE_HIDDEN_COLUMNS;

    /**
     * Specifies whether, for a merged region, to treat each individual cell as having the same value on import.
     *
     * @return {@code true} to give every cell of a merged region the region's value on import; defaults to {@link PxlConstants#DEFAULT_IMPORT_EACH_CELL_OF_MERGED_REGION} ({@code false})
     */
    boolean importEachCellOfMergedRegion() default PxlConstants.DEFAULT_IMPORT_EACH_CELL_OF_MERGED_REGION;

    /**
     * Specifies the index of the row to use as the header on import.
     * (The default is the first row; when specified separately, set it as 1-based, and the value must be less than the value of importFirstDataRowIndex.)
     *
     * @return the 1-based header row index on import; defaults to {@link PxlConstants#DEFAULT_IMPORT_HEADER_ROW_INDEX} ({@code 0}, i.e. the first row)
     */
    int importHeaderRowIndex() default PxlConstants.DEFAULT_IMPORT_HEADER_ROW_INDEX;

    /**
     * Specifies the index of the starting row to use as data on import.
     * (The default is the second row; when specified separately, set it as 1-based, and the value must be greater than the value of importHeaderRowIndex and less than or equal to the value of importLastDataRowIndex.)
     *
     * @return the 1-based first data row index on import; defaults to {@link PxlConstants#DEFAULT_IMPORT_FIRST_DATA_ROW_INDEX} ({@code 0}, i.e. the second row)
     */
    int importFirstDataRowIndex() default PxlConstants.DEFAULT_IMPORT_FIRST_DATA_ROW_INDEX;

    /**
     * Specifies the index of the ending row to use as data on import.
     * (The default is the last row; when specified separately, set it as 1-based, and the value must be greater than or equal to the value of importFirstDataRowIndex.)
     *
     * @return the 1-based last data row index on import; defaults to {@link PxlConstants#DEFAULT_IMPORT_LAST_DATA_ROW_INDEX} ({@code 0}, i.e. the last row)
     */
    int importLastDataRowIndex() default PxlConstants.DEFAULT_IMPORT_LAST_DATA_ROW_INDEX;

    /**
     * Specifies the index of the starting column to use as data on import.
     * (The default is the first column; when specified separately, set it as 1-based, and the value must be less than or equal to the value of importLastDataColumnIndex.)
     *
     * @return the 1-based first data column index on import; defaults to {@link PxlConstants#DEFAULT_IMPORT_FIRST_DATA_COLUMN_INDEX} ({@code 0}, i.e. the first column)
     */
    int importFirstDataColumnIndex() default PxlConstants.DEFAULT_IMPORT_FIRST_DATA_COLUMN_INDEX;

    /**
     * Specifies the index of the ending column to use as data on import.
     * (The default is the last column; when specified separately, set it as 1-based, and the value must be greater than or equal to the value of importFirstDataColumnIndex.)
     *
     * @return the 1-based last data column index on import; defaults to {@link PxlConstants#DEFAULT_IMPORT_LAST_DATA_COLUMN_INDEX} ({@code 0}, i.e. the last column)
     */
    int importLastDataColumnIndex() default PxlConstants.DEFAULT_IMPORT_LAST_DATA_COLUMN_INDEX;

    /**
     * Specifies the Character Encoding Set of the CSV to import, for this sheet alone.
     * Ignored for an Excel source, where the whole workbook is one file.
     * <p>
     * A CSV workbook is read as one file per sheet, so each sheet may carry its own charset. Left blank, the sheet
     * falls back to the workbook's {@code importCsvCharset}; a runtime sheet option overrides both.
     *
     * @return the character encoding used to read this sheet's CSV; defaults to {@link PxlConstants#UNSPECIFIED_IMPORT_CSV_CHARSET} ({@code ""}, i.e. inherit the workbook charset)
     * @see PxlWorkbook#importCsvCharset()
     * @see <a href="https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html">Java supported encodings</a>
     */
    String importCsvCharset() default PxlConstants.UNSPECIFIED_IMPORT_CSV_CHARSET;

    /**
     * Specifies the Delimiter of the CSV to import, for this sheet alone.
     * Ignored for an Excel source, where the whole workbook is one file.
     * <p>
     * A CSV workbook is read as one file per sheet, so each sheet may carry its own delimiter. Left at NUL, the sheet
     * falls back to the workbook's {@code importCsvDelimiter}; a runtime sheet option overrides both.
     *
     * @return the CSV field delimiter used to read this sheet; defaults to {@link PxlConstants#UNSPECIFIED_IMPORT_CSV_DELIMITER} ({@code '\0'}, i.e. inherit the workbook delimiter)
     * @see PxlWorkbook#importCsvDelimiter()
     */
    char importCsvDelimiter() default PxlConstants.UNSPECIFIED_IMPORT_CSV_DELIMITER;

    /**
     * Specifies whether to export.
     *
     * @return {@code true} to write this sheet on export; {@code true} by default
     */
    boolean exportEnabled() default true;

    /**
     * Specifies whether to export a sample.
     *
     * @return {@code true} to include this sheet in a sample export; {@code true} by default
     */
    boolean exportSampleEnabled() default true;

    /**
     * Specifies whether, on export, to override a superclass field that uses the same sheet name, if one exists.
     * Names differing only in case count as the same sheet name here, since a workbook cannot hold both.
     *
     * @return {@code true} to override a same-named superclass sheet on export; defaults to {@link PxlConstants#DEFAULT_EXPORT_OVERRIDE_SUPER_CLASS_SHEET} ({@code false})
     */
    boolean exportOverrideSuperClassSheet() default PxlConstants.DEFAULT_EXPORT_OVERRIDE_SUPER_CLASS_SHEET;

    /**
     * Specifies the height of rows within the sheet on export.
     *
     * @return the row height in points on export; defaults to {@link PxlConstants#DEFAULT_EXPORT_ROW_HEIGHT_IN_POINTS} ({@code -1}, i.e. unset)
     */
    float exportRowHeightInPoints() default PxlConstants.DEFAULT_EXPORT_ROW_HEIGHT_IN_POINTS;

    /**
     * Specifies the ordering between sheets on export. (in alphabetical order)
     *
     * @return the alphabetical sort key that orders sheets on export; empty ({@code ""}) by default
     */
    String exportOrder() default "";

    /**
     * Specifies the name of the field by which to group and split into multiple sheets on export.
     *
     * @return the field name used to group rows and split them into multiple sheets on export; empty ({@code ""}, no grouping) by default
     */
    String exportGroupingFieldName() default "";

    /**
     * Specifies the index of the row to use as the header on export.
     * (The default is the first row; when specified separately, set it as 1-based, and the value must be less than the value of exportFirstDataRowIndex.)
     *
     * @return the 1-based header row index on export; defaults to {@link PxlConstants#DEFAULT_EXPORT_HEADER_ROW_INDEX} ({@code 0}, i.e. the first row)
     */
    int exportHeaderRowIndex() default PxlConstants.DEFAULT_EXPORT_HEADER_ROW_INDEX;

    /**
     * Specifies the index of the starting row to use as data on export.
     * (The default is the second row; when specified separately, set it as 1-based, and the value must be greater than the value of exportHeaderRowIndex and less than or equal to the value of exportLastDataRowIndex.)
     *
     * @return the 1-based first data row index on export; defaults to {@link PxlConstants#DEFAULT_EXPORT_FIRST_DATA_ROW_INDEX} ({@code 0}, i.e. the second row)
     */
    int exportFirstDataRowIndex() default PxlConstants.DEFAULT_EXPORT_FIRST_DATA_ROW_INDEX;

    /**
     * Specifies the index of the ending row to use as data on export.
     * (The default is the last row; when specified separately, set it as 1-based, and the value must be greater than or equal to the value of exportFirstDataRowIndex.)
     *
     * @return the 1-based last data row index on export; defaults to {@link PxlConstants#DEFAULT_EXPORT_LAST_DATA_ROW_INDEX} ({@code 0}, i.e. the last row)
     */
    int exportLastDataRowIndex() default PxlConstants.DEFAULT_EXPORT_LAST_DATA_ROW_INDEX;

    /**
     * Specifies the index of the starting column to use as data on export.
     * (The default is the first column; when specified separately, set it as 1-based, and the value must be less than or equal to the value of exportLastDataColumnIndex.)
     *
     * @return the 1-based first data column index on export; defaults to {@link PxlConstants#DEFAULT_EXPORT_FIRST_DATA_COLUMN_INDEX} ({@code 0}, i.e. the first column)
     */
    int exportFirstDataColumnIndex() default PxlConstants.DEFAULT_EXPORT_FIRST_DATA_COLUMN_INDEX;

    /**
     * Specifies the index of the ending column to use as data on export.
     * (The default is the last column; when specified separately, set it as 1-based, and the value must be greater than or equal to the value of exportFirstDataColumnIndex.)
     *
     * @return the 1-based last data column index on export; defaults to {@link PxlConstants#DEFAULT_EXPORT_LAST_DATA_COLUMN_INDEX} ({@code 0}, i.e. the last column)
     */
    int exportLastDataColumnIndex() default PxlConstants.DEFAULT_EXPORT_LAST_DATA_COLUMN_INDEX;

    /**
     * Specifies whether to export when the data list is Null.
     *
     * @return {@code true} to still emit the sheet when the data list is {@code null}; defaults to {@link PxlConstants#DEFAULT_EXPORT_IF_NULL} ({@code false})
     */
    boolean exportIfNull() default PxlConstants.DEFAULT_EXPORT_IF_NULL;

    /**
     * Specifies whether to export when the data list is empty.
     *
     * @return {@code true} to still emit the sheet when the data list is empty; defaults to {@link PxlConstants#DEFAULT_EXPORT_IF_EMPTY} ({@code true})
     */
    boolean exportIfEmpty() default PxlConstants.DEFAULT_EXPORT_IF_EMPTY;

    /**
     * Specifies whether to apply a filter on export.
     *
     * @return {@code true} to apply an auto-filter to the columns on export; defaults to {@link PxlConstants#DEFAULT_EXPORT_COLUMN_FILTER} ({@code false})
     */
    boolean exportColumnFilter() default PxlConstants.DEFAULT_EXPORT_COLUMN_FILTER;

    /**
     * Specifies the style to apply to a required header cell on export.
     *
     * @return the styler for this sheet's required header cells on export; {@link PxlStyler} (unset) by default, deferring to the workbook/built-in styler
     */
    Class<? extends PxlStyler> exportSheetRequiredHeaderCellStyler() default PxlStyler.class;

    /**
     * Specifies the style to apply to an optional header cell on export.
     *
     * @return the styler for this sheet's optional header cells on export; {@link PxlStyler} (unset) by default, deferring to the workbook/built-in styler
     */
    Class<? extends PxlStyler> exportSheetOptionalHeaderCellStyler() default PxlStyler.class;

    /**
     * Specifies the style to apply to a Data Cell on export.
     *
     * @return the styler for this sheet's data cells on export; {@link PxlStyler} (unset) by default, deferring to the workbook/built-in styler
     */
    Class<? extends PxlStyler> exportSheetDataCellStyler() default PxlStyler.class;

}
