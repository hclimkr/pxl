package io.github.hclimkr.pxl.annotation;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.PxlFileFormat;
import io.github.hclimkr.pxl.styler.PxlStyler;

import java.lang.annotation.*;

/**
 * Annotation for specifying properties on an Excel workbook
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
@Documented
public @interface PxlWorkbook {

    /**
     * Specifies the Password used to remove document protection on import.
     *
     * @return the password used to open a protected document on import; defaults to {@link PxlConstants#DEFAULT_IMPORT_PASSWORD} ({@code ""}, no password)
     */
    String importPassword() default PxlConstants.DEFAULT_IMPORT_PASSWORD;

    /**
     * Specifies whether to validate the data to be imported.
     *
     * @return {@code true} to bean-validate imported data; defaults to {@link PxlConstants#DEFAULT_IMPORT_DATA_VALIDATION} ({@code true})
     */
    boolean importDataValidation() default PxlConstants.DEFAULT_IMPORT_DATA_VALIDATION;

    /**
     * Specifies whether to use the Stream Reader on import. (Works only with XSSF-format Excel files.)
     * https://github.com/pjfanning/excel-streaming-reader
     *
     * @return {@code true} to read XSSF files with the streaming reader; defaults to {@link PxlConstants#DEFAULT_IMPORT_USING_STREAM_READER} ({@code false})
     */
    boolean importUsingStreamReader() default PxlConstants.DEFAULT_IMPORT_USING_STREAM_READER;

    /**
     * Specifies the value of RowCacheSize when importing with the Stream Reader.
     *
     * @return the streaming reader row-cache size (number of rows kept in memory); defaults to {@link PxlConstants#DEFAULT_IMPORT_STREAM_READER_ROW_CACHE_SIZE} ({@code 100})
     */
    int importStreamReaderRowCacheSize() default PxlConstants.DEFAULT_IMPORT_STREAM_READER_ROW_CACHE_SIZE;

    /**
     * Specifies the value of BufferSize when importing with the Stream Reader.
     *
     * @return the streaming reader buffer size in bytes read from the input resource; defaults to {@link PxlConstants#DEFAULT_IMPORT_STREAM_READER_BUFFER_SIZE} ({@code 4096})
     */
    int importStreamReaderBufferSize() default PxlConstants.DEFAULT_IMPORT_STREAM_READER_BUFFER_SIZE;

    /**
     * Specifies the Character Encoding Set of the CSV to import.
     * https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html
     *
     * @return the character encoding used to read a CSV; defaults to {@link PxlConstants#DEFAULT_IMPORT_CSV_CHARSET} ({@code "UTF-8"})
     */
    String importCsvCharset() default PxlConstants.DEFAULT_IMPORT_CSV_CHARSET;

    /**
     * Specifies the Delimiter of the CSV to import.
     *
     * @return the CSV field delimiter on import; defaults to {@link PxlConstants#DEFAULT_IMPORT_CSV_DELIMITER} ({@code ','})
     */

    char importCsvDelimiter() default PxlConstants.DEFAULT_IMPORT_CSV_DELIMITER;

    /**
     * Specifies the BaseName of the Resource Bundle for internationalization support on import.
     * <p>
     * Once a base name is set, {@link PxlSheet#name()} and {@link PxlColumn#name()} — together with the sheet and
     * column names an import option overrides them with — are read as keys of that bundle, and sheet/header matching
     * runs against the translations. A key the bundle does not carry is matched as it stands.
     *
     * @return the resource-bundle base name for import i18n; defaults to {@link PxlConstants#DEFAULT_IMPORT_I18N_BASE_NAME} ({@code ""}, which disables i18n)
     */
    String importI18nBaseName() default PxlConstants.DEFAULT_IMPORT_I18N_BASE_NAME;

    /**
     * Specifies the Language of the Resource Bundle for internationalization support on import.
     *
     * @return the resource-bundle language for import i18n; defaults to {@link PxlConstants#DEFAULT_IMPORT_I18N_LANGUAGE} ({@code "en"})
     */
    String importI18nLanguage() default PxlConstants.DEFAULT_IMPORT_I18N_LANGUAGE;

    /**
     * Specifies the Country of the Resource Bundle for internationalization support on import.
     *
     * @return the resource-bundle country for import i18n; defaults to {@link PxlConstants#DEFAULT_IMPORT_I18N_COUNTRY} ({@code ""})
     */
    String importI18nCountry() default PxlConstants.DEFAULT_IMPORT_I18N_COUNTRY;

    /**
     * Specifies the format of the Excel workbook to be exported.
     *
     * @return the workbook file format produced on export; defaults to {@link PxlFileFormat#XSSF}
     */
    PxlFileFormat exportFileFormat() default PxlFileFormat.XSSF;

    /**
     * Specifies the Password used to protect the document on export.
     *
     * @return the password used to encrypt/protect the document on export; defaults to {@link PxlConstants#DEFAULT_EXPORT_PASSWORD} ({@code ""}, no protection)
     */
    String exportPassword() default PxlConstants.DEFAULT_EXPORT_PASSWORD;

    /**
     * Specifies whether to validate the data to be exported.
     *
     * @return {@code true} to bean-validate exported data; defaults to {@link PxlConstants#DEFAULT_EXPORT_DATA_VALIDATION} ({@code true})
     */
    boolean exportDataValidation() default PxlConstants.DEFAULT_EXPORT_DATA_VALIDATION;

    /**
     * Specifies the value of rowAccessWindowSize used when exporting as SXSSF.
     *
     * @return the SXSSF row-access window size (rows kept in memory) used on export; defaults to {@link PxlConstants#DEFAULT_EXPORT_SXSSF_ROW_ACCESS_WINDOW_SIZE}
     */
    int exportSXSSFRowAccessWindowSize() default PxlConstants.DEFAULT_EXPORT_SXSSF_ROW_ACCESS_WINDOW_SIZE;

    /**
     * Specifies the style to apply to a required header cell on export.
     *
     * @return the styler for required header cells on export; {@link PxlStyler} (unset) by default, deferring to the built-in default styler
     */
    Class<? extends PxlStyler> exportWorkbookRequiredHeaderCellStyler() default PxlStyler.class;

    /**
     * Specifies the style to apply to an optional header cell on export.
     *
     * @return the styler for optional header cells on export; {@link PxlStyler} (unset) by default, deferring to the built-in default styler
     */
    Class<? extends PxlStyler> exportWorkbookOptionalHeaderCellStyler() default PxlStyler.class;

    /**
     * Specifies the style to apply to a Data Cell on export.
     *
     * @return the styler for data cells on export; {@link PxlStyler} (unset) by default, deferring to the built-in default styler
     */
    Class<? extends PxlStyler> exportWorkbookDataCellStyler() default PxlStyler.class;

    /**
     * Specifies the BaseName of the Resource Bundle for internationalization support on export.
     * <p>
     * Once a base name is set, {@link PxlSheet#name()} and {@link PxlColumn#name()} — together with the sheet and
     * column names an export option overrides them with — are read as keys of that bundle and written as their
     * translations, as are the cell values described on {@link PxlColumn#exportSample()} and
     * {@link PxlColumn#exportOptionItems()}. A key the bundle does not carry is written as it stands.
     *
     * @return the resource-bundle base name for export i18n; defaults to {@link PxlConstants#DEFAULT_EXPORT_I18N_BASE_NAME} ({@code ""}, which disables i18n)
     */
    String exportI18nBaseName() default PxlConstants.DEFAULT_EXPORT_I18N_BASE_NAME;

    /**
     * Specifies the Language of the Resource Bundle for internationalization support on export.
     *
     * @return the resource-bundle language for export i18n; defaults to {@link PxlConstants#DEFAULT_EXPORT_I18N_LANGUAGE} ({@code "en"})
     */
    String exportI18nLanguage() default PxlConstants.DEFAULT_EXPORT_I18N_LANGUAGE;

    /**
     * Specifies the Country of the Resource Bundle for internationalization support on export.
     *
     * @return the resource-bundle country for export i18n; defaults to {@link PxlConstants#DEFAULT_EXPORT_I18N_COUNTRY} ({@code ""})
     */
    String exportI18nCountry() default PxlConstants.DEFAULT_EXPORT_I18N_COUNTRY;

}
