package io.github.hclimkr.pxl.builder;

import io.github.hclimkr.pxl.Pxl;
import io.github.hclimkr.pxl.exception.*;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.core.PxlCoreCsvImporter;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;
import org.apache.commons.io.IOUtils;

import javax.validation.Validator;
import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.*;

/**
 * CSV import builder. Created via {@link Pxl#importCsv()}.
 *
 * <p>The parse target is configured with {@link #workbook(Class)} (workbook form) or {@link #sheet(Class)}
 * (sheet form), each returning a {@link Source} whose source-terminal methods ({@code fromFile(...)} /
 * {@code fromFiles(...)} / {@code fromStream(...)} / {@code fromStreams(...)}) actually open the source and parse.
 * The workbook form parses multiple CSVs grouped by sheet, while the sheet form supports a single CSV only.</p>
 *
 * <p>Common settings (workbook name and override option) are provided by {@link PxlAbstractImportBuilder}. The same
 * settings are also available on {@link Source}, so {@code workbookName(...)}/{@code override(...)} may be chained
 * either before or after {@code workbook(...)}/{@code sheet(...)} (the value set last wins).</p>
 *
 * <p>Example: {@code List<Employee> rows = pxl.importCsv().sheet(Employee.class).fromFile(file);}</p>
 */
public final class PxlCsvImportBuilder extends PxlAbstractImportBuilder {

    /**
     * Creates a CSV import builder with the given validator.
     *
     * @param validator the bean-validation validator, or {@code null} when bean validation is disabled
     */
    public PxlCsvImportBuilder(final Validator validator) {

        super(validator);
    }

    /**
     * Specifies the workbook name. (For setting the name field in the workbook form; optional)
     *
     * <p>CSV import has no file-name fallback for it — a CSV file name names its <em>sheet</em> — so left unset, the
     * {@code @PxlWorkbookName} field stays untouched.</p>
     *
     * @param workbookName the workbook name, or {@code null} to leave the name field untouched
     * @return this builder
     */
    public PxlCsvImportBuilder workbookName(@Nullable final String workbookName) {

        this.workbookName = workbookName;
        return this;
    }

    /**
     * Overrides annotation-declared values with the given import option. (Optional)
     *
     * @param option the import option, or {@code null}
     * @return this builder
     */
    public PxlCsvImportBuilder override(@Nullable final PxlImportWorkbookOption option) {

        this.option = option;
        return this;
    }

    /**
     * Configures parsing multiple CSVs grouped by sheet into a {@code @PxlWorkbook} workbook object. Specify the
     * source and run the parse with the returned {@link Source}.
     *
     * @param workbookClass the workbook class
     * @param <W>           the workbook type
     * @return the source-terminal step returning the parsed workbook object
     * @throws PxlNullPointerException if {@code workbookClass} is {@code null}
     */
    public <W> Source<W> workbook(final Class<W> workbookClass)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(workbookClass, "workbookClass");

        return Source.forWorkbook(validator, workbookName, option, workbookClass);
    }

    /**
     * Configures parsing a single CSV into a list ({@link List}) of row objects. Specify the source and run the parse
     * with the returned {@link Source}. (There must be exactly one file/stream)
     *
     * @param rowClass the row class
     * @param <T>      the row type
     * @return the source-terminal step returning the parsed list of row objects
     * @throws PxlNullPointerException if {@code rowClass} is {@code null}
     */
    public <T> Source<List<T>> sheet(final Class<T> rowClass)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(rowClass, "rowClass");

        return Source.forSheet(validator, workbookName, option, List.class, rowClass);
    }

    /**
     * Configures parsing a single CSV into the specified collection type. Specify the source and run the parse with
     * the returned {@link Source}. (There must be exactly one file/stream)
     *
     * @param rowClass        the row class
     * @param collectionClass the return collection implementation/interface type (e.g. {@code List.class}, {@code Set.class})
     * @param <C>             the collection type
     * @return the source-terminal step returning the parsed collection of row objects
     * @throws PxlNullPointerException if {@code rowClass} or {@code collectionClass} is {@code null}
     */
    public <C extends Collection<?>> Source<C> sheet(final Class<?> rowClass,
                                                     final Class<C> collectionClass)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(rowClass, "rowClass");
        PxlAssertSupport.notNull(collectionClass, "collectionClass");

        return Source.forSheet(validator, workbookName, option, collectionClass, rowClass);
    }

    /**
     * Terminal source step for CSV import. Holds the parse spec (workbook or sheet) configured on the enclosing
     * {@link PxlCsvImportBuilder}; the source-terminal methods {@link #fromFile(File)} / {@link #fromFiles(List)} /
     * {@link #fromStream(String, InputStream)} / {@link #fromStreams(List, List)} open the source(s), run the parse,
     * and return the typed result {@code R}.
     *
     * <p>The workbook form parses multiple CSVs (files/streams) grouped by sheet; the sheet form supports a single CSV
     * only. Only obtained from {@code workbook(...)}/{@code sheet(...)} on the builder — hence a nested type.</p>
     *
     * <p>The source-terminal methods are the <strong>normalization boundary</strong>: they declare
     * {@code throws PxlException}, but since that type is abstract what actually surfaces is always a concrete
     * subtype — the matching one for a classified failure ({@link PxlArgumentException},
     * {@link PxlCellCodecException},
     * {@link PxlValidationException}, ...), and {@link PxlSystemException}
     * (carrying the original as its cause) for anything else, such as a CSV file that cannot be opened or read.</p>
     *
     * @param <R> the parsed result type (a workbook object, a {@code List<row>}, or a {@code Collection<row>})
     */
    public static final class Source<R> {

        private final Validator validator;
        private String workbookName;
        private PxlImportWorkbookOption option;

        /**
         * Workbook class (workbook form). Non-null selects the workbook form; null selects the sheet form.
         */
        private final Class<?> workbookClass;
        /**
         * Return collection type of the sheet form (e.g. {@code List.class}, {@code Set.class}).
         */
        private final Class<?> collectionClass;
        /**
         * Row class of the sheet form.
         */
        private final Class<?> rowClass;

        /**
         * Holds the parse spec. Exactly one of the workbook form ({@code workbookClass} non-null) or the sheet
         * form ({@code collectionClass}/{@code rowClass}) is populated.
         *
         * @param validator       the bean-validation validator, or {@code null} when bean validation is disabled
         * @param workbookName    the workbook name, or {@code null}
         * @param option          the import option, or {@code null}
         * @param workbookClass   the workbook class (workbook form), or {@code null} for the sheet form
         * @param collectionClass the return collection type (sheet form), or {@code null} for the workbook form
         * @param rowClass        the row class (sheet form), or {@code null} for the workbook form
         */
        private Source(final Validator validator,
                       final String workbookName,
                       final PxlImportWorkbookOption option,
                       final Class<?> workbookClass,
                       final Class<?> collectionClass,
                       final Class<?> rowClass) {

            this.validator = validator;
            this.workbookName = workbookName;
            this.option = option;
            this.workbookClass = workbookClass;
            this.collectionClass = collectionClass;
            this.rowClass = rowClass;
        }

        /**
         * Creates a workbook-form source step.
         *
         * @param validator     the bean-validation validator, or {@code null} when bean validation is disabled
         * @param workbookName  the workbook name, or {@code null}
         * @param option        the import option, or {@code null}
         * @param workbookClass the workbook class
         * @param <W>           the workbook type
         * @return a workbook-form source step
         */
        private static <W> Source<W> forWorkbook(final Validator validator,
                                                 final String workbookName,
                                                 final PxlImportWorkbookOption option,
                                                 final Class<W> workbookClass) {

            return new Source<>(validator, workbookName, option, workbookClass, null, null);
        }

        /**
         * Creates a sheet-form source step.
         *
         * @param validator       the bean-validation validator, or {@code null} when bean validation is disabled
         * @param workbookName    the workbook name, or {@code null}
         * @param option          the import option, or {@code null}
         * @param collectionClass the return collection type (e.g. {@code List.class}, {@code Set.class})
         * @param rowClass        the row class
         * @param <R>             the parsed result type
         * @return a sheet-form source step
         */
        private static <R> Source<R> forSheet(final Validator validator,
                                              final String workbookName,
                                              final PxlImportWorkbookOption option,
                                              final Class<?> collectionClass,
                                              final Class<?> rowClass) {

            return new Source<>(validator, workbookName, option, null, collectionClass, rowClass);
        }

        /**
         * Specifies the workbook name. (For setting the name field in the workbook form; optional)
         *
         * <p>Same as {@link PxlCsvImportBuilder#workbookName(String)}, so it may be chained either before or after
         * {@code workbook(...)}/{@code sheet(...)}; the value set last wins.</p>
         *
         * @param workbookName the workbook name, or {@code null}
         * @return this source step
         */
        public Source<R> workbookName(@Nullable final String workbookName) {

            this.workbookName = workbookName;
            return this;
        }

        /**
         * Overrides annotation-declared values with the given import option. (Optional)
         *
         * <p>Same as {@link PxlCsvImportBuilder#override(PxlImportWorkbookOption)}, so it may be chained either
         * before or after {@code workbook(...)}/{@code sheet(...)}; the value set last wins.</p>
         *
         * @param option the import option, or {@code null}
         * @return this source step
         */
        public Source<R> override(@Nullable final PxlImportWorkbookOption option) {

            this.option = option;
            return this;
        }

        /**
         * Opens the given single CSV file as the source, parses it, and returns the result.
         *
         * <p>The file is opened and closed internally, so the caller has nothing to close.</p>
         *
         * <p>One CSV file is one sheet, and its name comes from the file: the file name without its extension is
         * matched against the {@code @PxlSheet} name in the workbook form.</p>
         *
         * @param csvFile the CSV file
         * @return the parsed result
         * @throws PxlNullPointerException if {@code csvFile} is {@code null}
         * @throws PxlException            if the file cannot be opened or read, or if parsing fails
         */
        public R fromFile(final File csvFile)
                throws PxlException {

            PxlAssertSupport.notNull(csvFile, "csvFile");

            return fromFiles(Arrays.asList(csvFile));
        }

        /**
         * Opens the given CSV files as the source (for workbook-form parsing), parses them, and returns the result.
         *
         * <p>The files are opened and closed internally, so the caller has nothing to close.</p>
         *
         * <p>Each file is one sheet: the file name without its extension is matched against the {@code @PxlSheet}
         * names, so the files may be given in any order. Use {@link #fromStreams(List, List)} to name the sheets
         * explicitly instead.</p>
         *
         * @param csvFiles the CSV files
         * @return the parsed result
         * @throws PxlNullPointerException if {@code csvFiles} is {@code null}
         * @throws PxlArgumentException    if {@code csvFiles} is empty
         * @throws PxlException            if a file cannot be opened or read, or if parsing fails
         */
        public R fromFiles(final List<File> csvFiles)
                throws PxlException {

            PxlAssertSupport.notEmpty(csvFiles, "csvFiles");

            final List<InputStream> openedStreams = new ArrayList<>();

            try {
                final List<String> names = new ArrayList<>();
                for (final File csvFile : csvFiles) {
                    names.add(getNormalizedFileBaseName(csvFile));
                    openedStreams.add(new BufferedInputStream(new FileInputStream(csvFile)));
                }

                return parse(names, openedStreams);
            } catch (PxlException e) {
                throw e;
            } catch (Exception e) {
                throw new PxlSystemException(e);
            } finally {
                openedStreams.forEach(IOUtils::closeQuietly);
            }
        }

        /**
         * Uses the given single CSV input stream as the source, parses it, and returns the result.
         *
         * <p>The given stream is <strong>not closed</strong>; the caller retains ownership and is responsible for closing it.</p>
         *
         * @param csvName   the CSV name
         * @param csvStream the CSV input stream (not closed by this method)
         * @return the parsed result
         * @throws PxlNullPointerException if {@code csvName} or {@code csvStream} is {@code null}
         * @throws PxlException            if parsing fails
         */
        public R fromStream(final String csvName,
                            final InputStream csvStream)
                throws PxlException {

            PxlAssertSupport.notNull(csvName, "csvName");
            PxlAssertSupport.notNull(csvStream, "csvStream");

            return fromStreams(Arrays.asList(csvName), Arrays.asList(csvStream));
        }

        /**
         * Uses the given CSV input streams as the source (for workbook-form parsing), parses them, and returns the result.
         *
         * <p>The given streams are <strong>not closed</strong>; the caller retains ownership and is responsible for closing them.</p>
         *
         * @param csvNames   the CSV names
         * @param csvStreams the CSV input streams (not closed by this method)
         * @return the parsed result
         * @throws PxlNullPointerException if {@code csvNames} or {@code csvStreams} is {@code null}
         * @throws PxlArgumentException    if {@code csvNames} or {@code csvStreams} is empty
         * @throws PxlException            if parsing fails
         */
        public R fromStreams(final List<String> csvNames,
                             final List<InputStream> csvStreams)
                throws PxlException {

            PxlAssertSupport.notEmpty(csvNames, "csvNames");
            PxlAssertSupport.notEmpty(csvStreams, "csvStreams");

            try {
                return parse(csvNames, csvStreams);
            } catch (PxlException e) {
                throw e;
            } catch (Exception e) {
                throw new PxlSystemException(e);
            }
        }

        /**
         * Runs the configured parse over the given CSV names and streams and returns the typed result.
         *
         * @param names   the CSV names
         * @param streams the CSV input streams
         * @return the parsed result
         * @throws PxlArgumentException if the sheet form is given more than one source
         * @throws PxlException         if parsing fails
         */
        private R parse(final List<String> names, final List<InputStream> streams)
                throws PxlException {

            final Object result;
            if (Objects.nonNull(workbookClass)) {
                result = PxlCoreCsvImporter.parseCsv(workbookName, names, streams, workbookClass, option, validator);
            } else {
                if (names.size() != 1) {
                    throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.BUILDER_CSV_SINGLE_SOURCE_ONLY));
                }
                result = PxlCoreCsvImporter.parseCsv(names.get(0), streams.get(0), collectionClass, rowClass, option, validator);
            }

            @SuppressWarnings("unchecked") final R typed = (R) result;
            return typed;
        }

    }

}
