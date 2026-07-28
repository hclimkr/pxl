package io.github.hclimkr.pxl.builder;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.exception.*;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.core.PxlCoreExcelImporter;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlImportWorkbookMeta;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.internal.support.PxlWorkbookSupport;
import io.github.hclimkr.pxl.option.PxlImportSheetOption;
import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;
import io.github.hclimkr.pxl.util.PxlWorkbookUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;

import javax.validation.Validator;
import java.io.File;
import java.io.InputStream;
import java.util.*;

/**
 * Excel import builder. Created via {@link io.github.hclimkr.pxl.Pxl#importExcel()}.
 *
 * <p>The parse target is configured with {@link #workbook(Class)} (workbook form) or {@link #sheet(Class, String...)}
 * (sheet form), each returning a {@link Source} whose source-terminal methods ({@code fromFile(...)} /
 * {@code fromStream(...)}) actually open the source and parse.</p>
 *
 * <p>Common settings (workbook name and override option) are provided by {@link PxlAbstractImportBuilder}. The same
 * settings are also available on {@link Source}, so {@code workbookName(...)}/{@code override(...)} may be chained
 * either before or after {@code workbook(...)}/{@code sheet(...)} (the value set last wins).</p>
 *
 * <p>Example: {@code List<User> users = pxl.importExcel().sheet(User.class, "Users").fromFile(file);}</p>
 */
public final class PxlExcelImportBuilder extends PxlAbstractImportBuilder {

    /**
     * Creates an Excel import builder with the given validator.
     *
     * @param validator the bean-validation validator, or {@code null} when bean validation is disabled
     */
    public PxlExcelImportBuilder(final Validator validator) {

        super(validator);
    }

    /**
     * Specifies the workbook name. (For setting the name field in the workbook form; optional)
     *
     * @param workbookName the workbook name, or {@code null}
     * @return this builder
     */
    public PxlExcelImportBuilder workbookName(@Nullable final String workbookName) {

        this.workbookName = workbookName;
        return this;
    }

    /**
     * Overrides annotation-declared values with the given import option. (Optional)
     *
     * @param option the import option, or {@code null}
     * @return this builder
     */
    public PxlExcelImportBuilder override(@Nullable final PxlImportWorkbookOption option) {

        this.option = option;
        return this;
    }

    /**
     * Configures parsing a workbook object from the {@code @PxlWorkbook} class. Specify the source and run the parse
     * with the returned {@link Source}.
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
     * Configures parsing a list ({@code List}) of row objects from one of the candidate sheet names. Specify the source
     * and run the parse with the returned {@link Source}.
     *
     * @param rowClass            the row class
     * @param candidateSheetNames the candidate sheet names (one or more)
     * @param <T>                 the row type
     * @return the source-terminal step returning the parsed list of row objects
     * @throws PxlNullPointerException if {@code rowClass} or {@code candidateSheetNames} is {@code null}
     * @throws PxlArgumentException    if {@code candidateSheetNames} is empty
     */
    public <T> Source<List<T>> sheet(final Class<T> rowClass,
                                     final String... candidateSheetNames)
            throws PxlNullPointerException, PxlArgumentException {

        PxlAssertSupport.notNull(candidateSheetNames, "candidateSheetNames");

        return sheet(rowClass, Arrays.asList(candidateSheetNames));
    }

    /**
     * Configures parsing a list ({@code List}) of row objects from the candidate sheet name list. Specify the source
     * and run the parse with the returned {@link Source}.
     *
     * @param rowClass            the row class
     * @param candidateSheetNames the candidate sheet name list (one or more)
     * @param <T>                 the row type
     * @return the source-terminal step returning the parsed list of row objects
     * @throws PxlNullPointerException if {@code rowClass} or {@code candidateSheetNames} is {@code null}
     * @throws PxlArgumentException    if {@code candidateSheetNames} is empty
     */
    public <T> Source<List<T>> sheet(final Class<T> rowClass,
                                     final List<String> candidateSheetNames)
            throws PxlNullPointerException, PxlArgumentException {

        PxlAssertSupport.notNull(rowClass, "rowClass");
        PxlAssertSupport.notEmpty(candidateSheetNames, "candidateSheetNames");

        return Source.forSheet(validator, workbookName, option, List.class, rowClass, candidateSheetNames);
    }

    /**
     * Configures parsing, from one of the candidate sheet names, into the specified collection type. Specify the source
     * and run the parse with the returned {@link Source}.
     *
     * @param collectionClass     the return collection implementation/interface type (e.g. {@code List.class}, {@code Set.class})
     * @param rowClass            the row class
     * @param candidateSheetNames the candidate sheet names (one or more)
     * @param <C>                 the collection type
     * @return the source-terminal step returning the parsed collection of row objects
     * @throws PxlNullPointerException if {@code collectionClass}, {@code rowClass}, or {@code candidateSheetNames} is {@code null}
     * @throws PxlArgumentException    if {@code candidateSheetNames} is empty
     */
    public <C extends Collection<?>> Source<C> sheet(final Class<C> collectionClass,
                                                     final Class<?> rowClass,
                                                     final String... candidateSheetNames)
            throws PxlNullPointerException, PxlArgumentException {

        PxlAssertSupport.notNull(candidateSheetNames, "candidateSheetNames");

        return sheet(collectionClass, rowClass, Arrays.asList(candidateSheetNames));
    }

    /**
     * Configures parsing, from the candidate sheet name list, into the specified collection type. Specify the source
     * and run the parse with the returned {@link Source}.
     *
     * @param collectionClass     the return collection implementation/interface type (e.g. {@code List.class}, {@code Set.class})
     * @param rowClass            the row class
     * @param candidateSheetNames the candidate sheet name list (one or more)
     * @param <C>                 the collection type
     * @return the source-terminal step returning the parsed collection of row objects
     * @throws PxlNullPointerException if {@code collectionClass}, {@code rowClass}, or {@code candidateSheetNames} is {@code null}
     * @throws PxlArgumentException    if {@code candidateSheetNames} is empty
     */
    public <C extends Collection<?>> Source<C> sheet(final Class<C> collectionClass,
                                                     final Class<?> rowClass,
                                                     final List<String> candidateSheetNames)
            throws PxlNullPointerException, PxlArgumentException {

        PxlAssertSupport.notNull(collectionClass, "collectionClass");
        PxlAssertSupport.notNull(rowClass, "rowClass");
        PxlAssertSupport.notEmpty(candidateSheetNames, "candidateSheetNames");

        return Source.forSheet(validator, workbookName, option, collectionClass, rowClass, candidateSheetNames);
    }

    /**
     * Terminal source step for Excel import. Holds the parse spec (workbook or sheet) configured on the enclosing
     * {@link PxlExcelImportBuilder}; the source-terminal methods {@link #fromFile(File)} / {@link #fromStream(InputStream)}
     * open the source, run the parse, and return the typed result {@code R}.
     *
     * <p>Only obtained from {@code workbook(...)}/{@code sheet(...)} on the builder — hence a nested type.</p>
     *
     * <p>The source-terminal methods are the <strong>normalization boundary</strong>: they declare
     * {@code throws PxlException}, but since that type is abstract what actually surfaces is always a concrete
     * subtype — the matching one for a classified failure ({@link PxlIOException} when the workbook cannot be
     * opened, {@link PxlCellCodecException}, {@link PxlValidationException}, {@link PxlArgumentException}, ...),
     * and {@link PxlSystemException} (carrying the original as its cause) for anything else.</p>
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
         * Candidate sheet names of the sheet form.
         */
        private final List<String> candidateSheetNames;

        /**
         * Holds the parse spec. Exactly one of the workbook form ({@code workbookClass} non-null) or the sheet
         * form ({@code collectionClass}/{@code rowClass}/{@code candidateSheetNames}) is populated.
         *
         * @param validator           the bean-validation validator, or {@code null} when bean validation is disabled
         * @param workbookName        the workbook name, or {@code null}
         * @param option              the import option, or {@code null}
         * @param workbookClass       the workbook class (workbook form), or {@code null} for the sheet form
         * @param collectionClass     the return collection type (sheet form), or {@code null} for the workbook form
         * @param rowClass            the row class (sheet form), or {@code null} for the workbook form
         * @param candidateSheetNames the candidate sheet names (sheet form), or {@code null} for the workbook form
         */
        private Source(final Validator validator,
                       final String workbookName,
                       final PxlImportWorkbookOption option,
                       final Class<?> workbookClass,
                       final Class<?> collectionClass,
                       final Class<?> rowClass,
                       final List<String> candidateSheetNames) {

            this.validator = validator;
            this.workbookName = workbookName;
            this.option = option;
            this.workbookClass = workbookClass;
            this.collectionClass = collectionClass;
            this.rowClass = rowClass;
            this.candidateSheetNames = candidateSheetNames;
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

            return new Source<>(validator, workbookName, option, workbookClass, null, null, null);
        }

        /**
         * Creates a sheet-form source step.
         *
         * @param validator           the bean-validation validator, or {@code null} when bean validation is disabled
         * @param workbookName        the workbook name, or {@code null}
         * @param option              the import option, or {@code null}
         * @param collectionClass     the return collection type (e.g. {@code List.class}, {@code Set.class})
         * @param rowClass            the row class
         * @param candidateSheetNames the candidate sheet names
         * @param <R>                 the parsed result type
         * @return a sheet-form source step
         */
        private static <R> Source<R> forSheet(final Validator validator,
                                              final String workbookName,
                                              final PxlImportWorkbookOption option,
                                              final Class<?> collectionClass,
                                              final Class<?> rowClass,
                                              final List<String> candidateSheetNames) {

            return new Source<>(validator, workbookName, option, null, collectionClass, rowClass, candidateSheetNames);
        }

        /**
         * Specifies the workbook name. (For setting the name field in the workbook form; optional)
         *
         * <p>Same as {@link PxlExcelImportBuilder#workbookName(String)}, so it may be chained either before or after
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
         * <p>Same as {@link PxlExcelImportBuilder#override(PxlImportWorkbookOption)}, so it may be chained either
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
         * Opens the given Excel file as the source, parses it, and returns the result.
         *
         * <p>The file is opened and closed internally, so the caller has nothing to close.</p>
         *
         * @param excelFile the Excel file
         * @return the parsed result
         * @throws PxlNullPointerException if {@code excelFile} is {@code null}
         * @throws PxlIOException          if the file cannot be opened or read
         * @throws PxlException            if parsing fails
         */
        public R fromFile(final File excelFile)
                throws PxlException {

            PxlAssertSupport.notNull(excelFile, "excelFile");

            return parse(excelFile, null);
        }

        /**
         * Opens the given Excel input stream as the source, parses it, and returns the result.
         *
         * <p>The given stream is <strong>not closed</strong>; the caller retains ownership and is responsible for closing it.</p>
         *
         * @param excelStream the Excel input stream (not closed by this method)
         * @return the parsed result
         * @throws PxlNullPointerException if {@code excelStream} is {@code null}
         * @throws PxlIOException          if the stream cannot be read as a workbook
         * @throws PxlException            if parsing fails
         */
        public R fromStream(final InputStream excelStream)
                throws PxlException {

            PxlAssertSupport.notNull(excelStream, "excelStream");

            return parse(null, excelStream);
        }

        /**
         * Opens the workbook from the given file or stream, runs the configured parse, and returns the typed result.
         *
         * @param excelFile   the Excel file, or {@code null} when a stream is given
         * @param excelStream the Excel input stream, or {@code null} when a file is given
         * @return the parsed result
         * @throws PxlArgumentException if neither source is given
         * @throws PxlIOException       if the workbook cannot be opened
         * @throws PxlException         if parsing fails
         */
        private R parse(@Nullable final File excelFile, @Nullable final InputStream excelStream)
                throws PxlException {

            if (Objects.isNull(excelFile) && Objects.isNull(excelStream)) {
                throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.BUILDER_IMPORT_SOURCE_MISSING));
            }

            Workbook workbook = null;

            try {
                final PxlImportWorkbookMeta workbookMeta = PxlImportWorkbookMeta.makeImportWorkbookMeta(workbookClass, option);
                workbook = openWorkbook(workbookMeta, excelFile, excelStream);

                final Object result;
                if (Objects.nonNull(workbookClass)) {
                    final List<PxlImportSheetOption> sheetOptions = workbookMeta.getImportSheetOptions();
                    result = PxlCoreExcelImporter.parseExcel(workbookName, workbook, workbookClass, workbookMeta, sheetOptions, validator);
                } else {
                    final PxlImportSheetOption sheetOption = Optional.ofNullable(workbookMeta.getImportSheetOptions())
                            .flatMap(options -> options.stream()
                                    .filter(o -> StringUtils.equals(o.getFieldName(), PxlConstants.SHEET_FIELD_NAME_WILD_CARD))
                                    .findFirst())
                            .orElse(null);
                    result = PxlCoreExcelImporter.parseExcel(workbook, candidateSheetNames, collectionClass, rowClass, workbookMeta, sheetOption, validator);
                }

                @SuppressWarnings("unchecked") final R typed = (R) result;
                return typed;
            } catch (PxlException e) {
                throw e;
            } catch (Exception e) {
                throw new PxlSystemException(e);
            } finally {
                PxlWorkbookUtils.closeWorkbook(workbook);
            }
        }

        /**
         * Opens a POI workbook from the file when present, otherwise from the stream, per the resolved workbook meta.
         *
         * @param workbookMeta the resolved import workbook meta (selects XLS/XLSX/streaming)
         * @param excelFile    the Excel file, or {@code null} when a stream is given
         * @param excelStream  the Excel input stream (used when {@code excelFile} is {@code null})
         * @return the opened workbook
         * @throws PxlIOException if the workbook cannot be opened (a missing file, an unreadable or
         *                        password-protected container, an unsupported format, ...) — the underlying
         *                        failure is wrapped as the cause
         */
        private Workbook openWorkbook(final PxlImportWorkbookMeta workbookMeta,
                                      final File excelFile,
                                      final InputStream excelStream)
                throws PxlIOException {

            try {
                if (Objects.nonNull(excelFile)) {
                    return PxlWorkbookSupport.openWorkbook(excelFile, workbookMeta);
                }
                return PxlWorkbookSupport.openWorkbook(excelStream, workbookMeta);
            } catch (Exception e) {
                throw new PxlIOException(e);
            }
        }

    }

}
