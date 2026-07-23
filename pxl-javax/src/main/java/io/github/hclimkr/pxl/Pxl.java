package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.builder.PxlCsvImportBuilder;
import io.github.hclimkr.pxl.builder.PxlExcelExportBuilder;
import io.github.hclimkr.pxl.builder.PxlExcelImportBuilder;
import io.github.hclimkr.pxl.builder.PxlSampleExcelExportBuilder;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.support.PxlReflectionSupport;
import io.github.hclimkr.pxl.internal.support.PxlWorkbookSupport;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.validation.Validation;
import javax.validation.Validator;
import java.lang.reflect.Field;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

/**
 * Public entry point for Excel/CSV file binding.
 *
 * <p>All import/export operations are performed with fluent builders.</p>
 * <ul>
 *   <li>export: {@link #exportExcel()} — data → Excel workbook/file/stream</li>
 *   <li>sample: {@link #exportSampleExcel()} — class → Excel template with one sample data row (each column's exportSample value)</li>
 *   <li>import: {@link #importExcel()} / {@link #importCsv()} — configure the target, then supply the source → object</li>
 * </ul>
 *
 * <p>Example:</p>
 * <pre>{@code
 * Pxl pxl = new Pxl();
 * pxl.exportExcel().sheet("Users", users, User.class).override(opt).toFile(file);
 * List<User> loaded = pxl.importExcel().sheet(User.class, "Users").fromFile(file);
 * }</pre>
 */
public final class Pxl {

    private static final Logger LOGGER = LoggerFactory.getLogger(Pxl.class);

    private final Validator validator;

    /**
     * Creates a {@code Pxl} instance, initializing the default bean-validation {@link Validator} shared
     * by the import/export builders it creates.
     *
     * <p>Bean validation is <strong>optional</strong>: if no provider (e.g. hibernate-validator) or no EL
     * implementation (e.g. jakarta.el) is on the classpath, initialization is skipped gracefully — a warning
     * is logged and the {@link Validator} is left {@code null}, so import/export proceeds with bean validation
     * simply disabled (no exception is thrown). Add a provider and EL to enable it.</p>
     */
    public Pxl() {

        // Bean validation is optional. If neither a provider (NoProviderFoundException) nor an EL implementation
        // (ValidationException HV000183) is present, buildDefaultValidatorFactory() throws at construction time in
        // both cases, so we catch it here, disable validation only (validator=null), and continue. (Every validation
        // call site skips a null validator.)
        Validator resolvedValidator;
        try {
            resolvedValidator = Validation.buildDefaultValidatorFactory().getValidator();
        } catch (Exception e) {
            resolvedValidator = null;
            LOGGER.warn(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.LOG_BEAN_VALIDATION_INIT_FAILED, e.getMessage()));
        }

        this.validator = resolvedValidator;
    }

    /**
     * Creates a builder that exports data to an Excel workbook/file/stream.
     *
     * @return Excel export builder
     */
    public PxlExcelExportBuilder exportExcel() {

        return new PxlExcelExportBuilder(validator);
    }

    /**
     * Creates a builder that generates an Excel sample template from a class: a header row plus a single
     * sample data row filled from each column's {@code exportSample} value (not an empty template).
     *
     * @return Excel sample builder
     */
    public PxlSampleExcelExportBuilder exportSampleExcel() {

        return new PxlSampleExcelExportBuilder();
    }

    /**
     * Creates an Excel import builder. Configure the parse target with {@code workbook(...)}/{@code sheet(...)}, then
     * supply the source and run the parse with the returned source step's {@code fromFile(...)}/{@code fromStream(...)}.
     *
     * @return Excel import builder
     */
    public PxlExcelImportBuilder importExcel() {

        return new PxlExcelImportBuilder(validator);
    }

    /**
     * Creates a CSV import builder. Configure the parse target with {@code workbook(...)}/{@code sheet(...)}, then
     * supply the source and run the parse with the returned source step's
     * {@code fromFile(...)}/{@code fromFiles(...)}/{@code fromStream(...)}/{@code fromStreams(...)}.
     *
     * @return CSV import builder
     */
    public PxlCsvImportBuilder importCsv() {

        return new PxlCsvImportBuilder(validator);
    }

    /**
     * Overrides the locale used for the library's own diagnostic messages (exception text and the
     * location tag) process-wide.
     *
     * <p>This is independent of the per-workbook sheet/column name translation configured via
     * {@code @PxlWorkbook}: diagnostic messages are a library-owned, process-wide concern. When no override
     * is set the JVM default locale ({@link Locale#getDefault()}) is used; the base bundle is English and
     * Korean is available. Pass {@code null} (or call {@link #resetMessageLocale()}) to clear the override.</p>
     *
     * @param locale the locale for diagnostic messages; {@code null} reverts to the JVM default
     */
    public static void setMessageLocale(final Locale locale) {

        PxlI18nDiagnostic.setOverrideLocale(locale);
    }

    /**
     * Clears any locale override set via {@link #setMessageLocale(Locale)}, reverting diagnostic messages
     * to the JVM default locale ({@link Locale#getDefault()}).
     */
    public static void resetMessageLocale() {

        PxlI18nDiagnostic.setOverrideLocale(null);
    }

    /**
     * Finds and returns the name from a workbook object.
     *
     * @param workbookObject workbook object
     * @return workbook name (null if absent)
     */
    public static String getWorkbookNameFromWorkbookObject(final Object workbookObject) {

        if (Objects.isNull(workbookObject)) {
            return null;
        }

        String workbookName = null;

        final Field workbookNameField = PxlWorkbookSupport.getWorkbookNameField(workbookObject.getClass());
        if (Objects.nonNull(workbookNameField)) {
            try {
                workbookName = (String) PxlReflectionSupport.getFieldValue(workbookNameField, workbookObject);
            } catch (Exception ignored) {
            }
        }

        return workbookName;
    }

    /**
     * Finds and returns the file format from a workbook class.
     *
     * @param workbookClass workbook class
     * @return file format (default value if absent)
     */
    public static PxlFileFormat getWorkbookFileFormatFromWorkbookObject(final Class<?> workbookClass) {

        if (Objects.isNull(workbookClass)) {
            return PxlConstants.DEFAULT_EXPORT_FILE_FORMAT;
        }

        final PxlWorkbook workbookAnnotation = workbookClass.getAnnotation(PxlWorkbook.class);

        return Optional.ofNullable(workbookAnnotation)
                .map(PxlWorkbook::exportFileFormat)
                .orElse(PxlConstants.DEFAULT_EXPORT_FILE_FORMAT);
    }

}
