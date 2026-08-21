package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.builder.*;
import io.github.hclimkr.pxl.internal.constraint.PxlSheetCascadeSkippingResolver;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.validation.Configuration;
import javax.validation.Validation;
import javax.validation.Validator;
import java.util.Locale;

/**
 * Public entry point for Excel/CSV file binding.
 *
 * <p>All export/import operations are performed with fluent builders.</p>
 * <ul>
 *   <li>export: {@link #exportExcel()} / {@link #exportCsv()} - data -> workbook/file/stream</li>
 *   <li>sample: {@link #exportSampleExcel()} / {@link #exportSampleCsv()} - class -> template with one sample data row (each column's exportSample value)</li>
 *   <li>import: {@link #importExcel()} / {@link #importCsv()} - configure the target, then supply the source -> object</li>
 * </ul>
 *
 * <p>Example:</p>
 * <pre>{@code
 * Pxl pxl = new Pxl();
 * pxl.exportExcel().sheet(User.class, users, "Users").override(opt).toFile(file);
 * List<User> loaded = pxl.importExcel().sheet(User.class, "Users").fromFile(file);
 * }</pre>
 */
public final class Pxl {

    private static final Logger LOGGER = LoggerFactory.getLogger(Pxl.class);

    private final Validator validator;

    /**
     * Creates a {@link Pxl} instance, initializing the bean-validation {@link Validator} shared by the
     * export/import builders it creates.
     *
     * <p>Bean validation is <strong>optional</strong>: if no provider (e.g. hibernate-validator) or no EL
     * implementation (e.g. jakarta.el) is on the classpath, initialization is skipped gracefully - a warning
     * is logged and the {@link Validator} is left {@code null}, so export/import proceeds with bean validation
     * simply disabled (no exception is thrown). Add a provider and EL to enable it.</p>
     *
     * <p>The validator skips one kind of cascade: validating a workbook object does not descend into the rows of a
     * {@link PxlSheet} field, even when that field is also marked {@code @Valid}. The binder validates those rows
     * itself, one at a time and with the sheet name - and, on import, the row index - attached, so cascading would
     * repeat the same work and report any violation without a location. Sheet rows are therefore validated exactly
     * where the binder processes them: a sheet it skips ({@code exportEnabled = false} on export,
     * {@code importEnabled = false} on import) has its rows left unvalidated, {@code @Valid} or not. Constraints on
     * the collection itself ({@code @NotEmpty}, {@code @Size}, ...) and a {@code @Valid} on any other field are
     * unaffected.</p>
     */
    public Pxl() {

        // Bean validation is optional. If neither a provider (NoProviderFoundException) nor an EL implementation
        // (ValidationException HV000183) is present, building the factory throws at construction time in both
        // cases, so we catch it here, disable validation only (validator=null), and continue. (Every validation
        // call site skips a null validator.)
        Validator resolvedValidator;
        try {
            // resolvedValidator = Validation.buildDefaultValidatorFactory().getValidator();

            final Configuration<?> configuration = Validation.byDefaultProvider().configure();

            // The resolver decides per root type, and a row class has no @PxlSheet field, so the same validator
            // serves the row pass too - there it has nothing to skip.
            configuration.traversableResolver(new PxlSheetCascadeSkippingResolver(configuration.getDefaultTraversableResolver()));

            resolvedValidator = configuration.buildValidatorFactory().getValidator();
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
     * Creates a builder that exports data to a CSV file/stream.
     *
     * <p>CSV is one file per sheet, so the builder has the sheet form only and its terminals write a single
     * sheet.</p>
     *
     * @return CSV export builder
     */
    public PxlCsvExportBuilder exportCsv() {

        return new PxlCsvExportBuilder(validator);
    }

    /**
     * Creates a builder that generates a CSV sample template from a class: a header record plus a single sample
     * data record filled from each column's {@code exportSample} value (not an empty template).
     *
     * @return CSV sample builder
     */
    public PxlSampleCsvExportBuilder exportSampleCsv() {

        return new PxlSampleCsvExportBuilder();
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

}
