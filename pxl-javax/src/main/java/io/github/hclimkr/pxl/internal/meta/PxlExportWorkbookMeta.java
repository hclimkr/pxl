package io.github.hclimkr.pxl.internal.meta;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.exception.PxlDataException;
import io.github.hclimkr.pxl.exception.PxlI18nException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nContent;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.support.PxlReflectionSupport;
import io.github.hclimkr.pxl.internal.support.PxlWorkbookSupport;
import io.github.hclimkr.pxl.option.PxlExportSheetOption;
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import io.github.hclimkr.pxl.styler.PxlStyler;
import io.github.hclimkr.pxl.type.PxlExcelEngine;
import io.github.hclimkr.pxl.type.PxlFileFormat;
import io.github.hclimkr.pxl.type.PxlOptionalBoolean;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import io.github.hclimkr.pxl.util.PxlMiscUtils;
import io.github.hclimkr.pxl.util.PxlWorkbookUtils;
import lombok.Getter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.*;

/**
 * Workbook export metadata, resolved for an Excel or a CSV destination alike; which one it is shows in
 * {@code exportFileFormat}, and a CSV destination is the one that carries no POI workbook.
 */
@Getter
public final class PxlExportWorkbookMeta {

    private static final Logger LOGGER = LoggerFactory.getLogger(PxlExportWorkbookMeta.class);

    // null on the CSV export path, which writes no POI workbook at all.
    @Nullable
    private final Workbook workbook;

    // null on the CSV export path, for the same reason as the workbook above.
    @Nullable
    private final FormulaEvaluator formulaEvaluator;

    // Map for reusing stylers
    private final Map<Class<? extends PxlStyler>, CellStyle> cellStyleMap;

    // Styler classes that failed to apply. Prevents workbook style count explosion from retrying createCellStyle for each cell.
    private final Set<Class<? extends PxlStyler>> failedCellStyleSet;

    // Cache of quote-prefix data cell styles (baseStyleIndex -> quotePrefixedStyle).
    // Creating a new CellStyle for each cell could exceed the workbook style count limit (XLS ~4000 / XLSX ~64000),
    // so create and reuse only one per base style.
    private final Map<Integer, CellStyle> quotePrefixedStyleCache = new HashMap<>();

    // Cache of styles combining (base style + date display format) when exporting a date/time as a Numeric cell.
    // The key is "baseStyleIndex|excelFormatCode", and for the same reason as the quote-prefix cache, only one is created and reused per combination.
    private final Map<String, CellStyle> dateFormattedStyleCache = new HashMap<>();

    // The POI engine that writes this workbook. The physical format below is derived from it, so the two never disagree.
    private final PxlExcelEngine exportExcelEngine;

    // Physical format of the output; the source of the sheet/row/column limits the binder enforces.
    private final PxlFileFormat exportFileFormat;

    private final String exportPassword;

    private final boolean exportDataValidation;

    private final int exportSXSSFRowAccessWindowSize;

    // CSV only; the sheet meta resolves its own value against this one.
    private final String exportCsvCharset;

    // CSV only; resolved the same way as the charset above.
    private final char exportCsvDelimiter;

    // CSV only; whether a byte order mark precedes the output.
    private final boolean exportCsvBom;

    private final Class<? extends PxlStyler> exportWorkbookRequiredHeaderCellStyler;

    private final Class<? extends PxlStyler> exportWorkbookOptionalHeaderCellStyler;

    private final Class<? extends PxlStyler> exportWorkbookDataCellStyler;

    private final ResourceBundle exportResourceBundle;

    private final List<PxlExportSheetOption> exportSheetOptions;

    private final List<PxlExportSheetMeta> exportSheetMetas;

    /**
     * Creates the resolved export metadata for a workbook, storing the given POI workbook and formula evaluator,
     * the merged option/annotation values, and initializing the cell-style reuse caches.
     *
     * @param workbook                               the newly created POI workbook; {@code null} on the CSV export path
     * @param formulaEvaluator                       the formula evaluator for the workbook; {@code null} on the CSV export path
     * @param exportExcelEngine                      the resolved POI engine writing the workbook; {@code null} on the CSV export path
     * @param exportFileFormat                       the physical output file format the engine produces
     * @param exportPassword                         the password for encrypted output; may be {@code null}
     * @param exportDataValidation                   whether bean validation is applied
     * @param exportSXSSFRowAccessWindowSize         the SXSSF streaming row access window size
     * @param exportCsvCharset                       the resolved workbook-level CSV charset (CSV destinations only)
     * @param exportCsvDelimiter                     the resolved workbook-level CSV field delimiter (CSV destinations only)
     * @param exportCsvBom                           whether a byte order mark precedes the CSV output
     * @param exportWorkbookRequiredHeaderCellStyler the workbook-level header cell styler for required columns
     * @param exportWorkbookOptionalHeaderCellStyler the workbook-level header cell styler for optional columns
     * @param exportWorkbookDataCellStyler           the workbook-level data cell styler
     * @param exportResourceBundle                   the content i18n bundle for sheet/column name translation; may be {@code null}
     * @param exportSheetOptions                     the per-sheet export overrides
     * @param exportSheetMetas                       the (initially empty) sheet metadata list
     */
    private PxlExportWorkbookMeta(final Workbook workbook,
                                  final FormulaEvaluator formulaEvaluator,
                                  final PxlExcelEngine exportExcelEngine,
                                  final PxlFileFormat exportFileFormat,
                                  final String exportPassword,
                                  final boolean exportDataValidation,
                                  final int exportSXSSFRowAccessWindowSize,
                                  final String exportCsvCharset,
                                  final char exportCsvDelimiter,
                                  final boolean exportCsvBom,
                                  final Class<? extends PxlStyler> exportWorkbookRequiredHeaderCellStyler,
                                  final Class<? extends PxlStyler> exportWorkbookOptionalHeaderCellStyler,
                                  final Class<? extends PxlStyler> exportWorkbookDataCellStyler,
                                  final ResourceBundle exportResourceBundle,
                                  final List<PxlExportSheetOption> exportSheetOptions,
                                  final List<PxlExportSheetMeta> exportSheetMetas) {

        this.workbook = workbook;
        this.formulaEvaluator = formulaEvaluator;
        this.cellStyleMap = new HashMap<>();
        this.failedCellStyleSet = new HashSet<>();
        this.exportExcelEngine = exportExcelEngine;
        this.exportFileFormat = exportFileFormat;
        this.exportPassword = exportPassword;
        this.exportDataValidation = exportDataValidation;
        this.exportSXSSFRowAccessWindowSize = exportSXSSFRowAccessWindowSize;
        this.exportCsvCharset = exportCsvCharset;
        this.exportCsvDelimiter = exportCsvDelimiter;
        this.exportCsvBom = exportCsvBom;
        this.exportWorkbookRequiredHeaderCellStyler = exportWorkbookRequiredHeaderCellStyler;
        this.exportWorkbookOptionalHeaderCellStyler = exportWorkbookOptionalHeaderCellStyler;
        this.exportWorkbookDataCellStyler = exportWorkbookDataCellStyler;
        this.exportResourceBundle = exportResourceBundle;
        this.exportSheetOptions = exportSheetOptions;
        this.exportSheetMetas = exportSheetMetas;
    }

    /**
     * On export, collects the workbook metadata from the workbook option and the workbook class.
     * The workbook option takes precedence over the workbook class.
     * A new empty POI workbook (created by the resolved engine) and its formula evaluator are created here.
     *
     * @param workbookClass  the {@link PxlWorkbook}-annotated workbook class supplying annotation defaults; may be {@code null}
     * @param workbookOption runtime overrides taking precedence over the class annotation; may be {@code null}
     * @return the assembled export workbook metadata, holding the newly created workbook, formula evaluator, resolved engine/format/password/validation settings, cascaded stylers, i18n bundle and sheet options
     * @throws PxlDataException        if the workbook name field type is invalid
     * @throws PxlI18nException        if the export content i18n bundle cannot be found for the configured base name and locale
     * @throws PxlNullPointerException if the created POI workbook is unexpectedly absent
     */
    public static PxlExportWorkbookMeta makeExportWorkbookMeta(@Nullable final Class<?> workbookClass,
                                                               @Nullable final PxlExportWorkbookOption workbookOption)
            throws PxlDataException, PxlI18nException, PxlNullPointerException {

        return makeExportWorkbookMetaInternally(workbookClass, workbookOption, false);
    }

    /**
     * On CSV export, collects the workbook metadata the same way {@link #makeExportWorkbookMeta} does, except that
     * no POI workbook is created.
     *
     * <p>CSV has no POI writer, so the engine is left {@code null} and the file format is fixed to
     * {@link PxlFileFormat#CSV} - which is what makes the binder enforce the CSV sheet/row/column limits.
     * This is a complete state rather than a deficient one: {@code getWorkbook()} and
     * {@code getFormulaEvaluator()} simply answer {@code null} on this path.</p>
     *
     * @param workbookClass  the {@link PxlWorkbook}-annotated workbook class supplying annotation defaults; may be {@code null}
     * @param workbookOption runtime overrides taking precedence over the class annotation; may be {@code null}
     * @return the assembled export workbook metadata, carrying no POI workbook and the CSV file format
     * @throws PxlDataException        if the workbook name field type is invalid
     * @throws PxlI18nException        if the export content i18n bundle cannot be found for the configured base name and locale
     * @throws PxlNullPointerException never on this path - CSV creates no workbook and no formula evaluator
     */
    public static PxlExportWorkbookMeta makeExportWorkbookMetaForCsv(@Nullable final Class<?> workbookClass,
                                                                     @Nullable final PxlExportWorkbookOption workbookOption)
            throws PxlDataException, PxlI18nException, PxlNullPointerException {

        return makeExportWorkbookMetaInternally(workbookClass, workbookOption, true);
    }

    /**
     * Merges the option and annotation values into the workbook metadata, creating the POI workbook only for an
     * Excel destination.
     *
     * @param workbookClass  the {@link PxlWorkbook}-annotated workbook class supplying annotation defaults; may be {@code null}
     * @param workbookOption runtime overrides taking precedence over the class annotation; may be {@code null}
     * @param forCsv         {@code true} to assemble metadata for a CSV destination (no POI workbook, CSV file format)
     * @return the assembled export workbook metadata
     * @throws PxlDataException        if the workbook name field type is invalid
     * @throws PxlI18nException        if the export content i18n bundle cannot be found for the configured base name and locale
     * @throws PxlNullPointerException if the created POI workbook is unexpectedly absent
     */
    private static PxlExportWorkbookMeta makeExportWorkbookMetaInternally(@Nullable final Class<?> workbookClass,
                                                                          @Nullable final PxlExportWorkbookOption workbookOption,
                                                                          final boolean forCsv)
            throws PxlDataException, PxlI18nException, PxlNullPointerException {

        PxlWorkbookSupport.validateWorkbookNameFieldType(workbookClass);

        final PxlWorkbook workbookAnnotation = Optional.ofNullable(workbookClass)
                .map(c -> c.getAnnotation(PxlWorkbook.class))
                .orElse(null);

        // CSV has no POI writer, so the engine stays null there and the format is stated directly instead.
        final PxlExcelEngine exportExcelEngine = forCsv ? null : Optional.ofNullable(workbookOption)
                .flatMap(option -> Optional.ofNullable(option.getExportExcelEngine()))
                .orElseGet(() -> Optional.ofNullable(workbookAnnotation)
                        .map(PxlWorkbook::exportExcelEngine)
                        .orElse(PxlConstants.DEFAULT_EXPORT_EXCEL_ENGINE));

        // For Excel the physical format is not declared separately: it is whatever the chosen engine writes.
        final PxlFileFormat exportFileFormat = forCsv ? PxlFileFormat.CSV : exportExcelEngine.getFileFormat();

        final String exportPassword = Optional.ofNullable(workbookOption)
                .flatMap(option -> Optional.ofNullable(option.getExportPassword()))
                .orElseGet(() -> Optional.ofNullable(workbookAnnotation)
                        .map(PxlWorkbook::exportPassword)
                        .orElse(PxlConstants.DEFAULT_EXPORT_PASSWORD));

        final boolean exportDataValidation = Optional.ofNullable(workbookOption)
                .flatMap(option -> Optional.ofNullable(option.getExportDataValidation()))
                .orElseGet(() -> Optional.ofNullable(workbookAnnotation)
                        .map(PxlWorkbook::exportDataValidation)
                        .orElse(PxlConstants.DEFAULT_EXPORT_DATA_VALIDATION));

        final int exportSXSSFRowAccessWindowSize = Optional.ofNullable(workbookOption)
                .flatMap(option -> Optional.ofNullable(option.getExportSXSSFRowAccessWindowSize()))
                .orElseGet(() -> Optional.ofNullable(workbookAnnotation)
                        .map(PxlWorkbook::exportSXSSFRowAccessWindowSize)
                        .orElse(PxlConstants.DEFAULT_EXPORT_SXSSF_ROW_ACCESS_WINDOW_SIZE));

        final String exportCsvCharset = Optional.ofNullable(workbookOption)
                .flatMap(option -> Optional.ofNullable(option.getExportCsvCharset()))
                .filter(StringUtils::isNotBlank)
                .orElseGet(() -> Optional.ofNullable(workbookAnnotation)
                        .map(PxlWorkbook::exportCsvCharset)
                        .filter(StringUtils::isNotBlank)
                        .orElse(PxlConstants.DEFAULT_EXPORT_CSV_CHARSET));

        final char exportCsvDelimiter = Optional.ofNullable(workbookOption)
                .flatMap(option -> Optional.ofNullable(option.getExportCsvDelimiter()))
                .filter(delimiter -> delimiter != PxlConstants.UNSPECIFIED_EXPORT_CSV_DELIMITER)
                .orElseGet(() -> Optional.ofNullable(workbookAnnotation)
                        .map(PxlWorkbook::exportCsvDelimiter)
                        .filter(delimiter -> delimiter != PxlConstants.UNSPECIFIED_EXPORT_CSV_DELIMITER)
                        .orElse(PxlConstants.DEFAULT_EXPORT_CSV_DELIMITER));

        final boolean exportCsvBom = Optional.ofNullable(workbookOption)
                .flatMap(option -> Optional.ofNullable(option.getExportCsvBom()))
                .orElseGet(() -> Optional.ofNullable(workbookAnnotation)
                        .map(PxlWorkbook::exportCsvBom)
                        .filter(PxlOptionalBoolean::isSpecified)
                        .map(PxlOptionalBoolean::toBoolean)
                        .orElse(PxlConstants.DEFAULT_EXPORT_CSV_BOM));

        Class<? extends PxlStyler> exportWorkbookRequiredHeaderCellStyler = null;
        if (Objects.nonNull(workbookOption) && Objects.nonNull(workbookOption.getExportWorkbookRequiredHeaderCellStyler())) {
            exportWorkbookRequiredHeaderCellStyler = workbookOption.getExportWorkbookRequiredHeaderCellStyler();
        } else if (Objects.nonNull(workbookAnnotation) && Objects.nonNull(workbookAnnotation.exportWorkbookRequiredHeaderCellStyler())) {
            exportWorkbookRequiredHeaderCellStyler = workbookAnnotation.exportWorkbookRequiredHeaderCellStyler();
        }
        if (!PxlMiscUtils.isEffectiveCellStylerClass(exportWorkbookRequiredHeaderCellStyler)) {
            exportWorkbookRequiredHeaderCellStyler = PxlConstants.DEFAULT_EXPORT_WORKBOOK_REQUIRED_HEADER_CELL_STYLER;
        }

        Class<? extends PxlStyler> exportWorkbookOptionalHeaderCellStyler = null;
        if (Objects.nonNull(workbookOption) && Objects.nonNull(workbookOption.getExportWorkbookOptionalHeaderCellStyler())) {
            exportWorkbookOptionalHeaderCellStyler = workbookOption.getExportWorkbookOptionalHeaderCellStyler();
        } else if (Objects.nonNull(workbookAnnotation) && Objects.nonNull(workbookAnnotation.exportWorkbookOptionalHeaderCellStyler())) {
            exportWorkbookOptionalHeaderCellStyler = workbookAnnotation.exportWorkbookOptionalHeaderCellStyler();
        }
        if (!PxlMiscUtils.isEffectiveCellStylerClass(exportWorkbookOptionalHeaderCellStyler)) {
            exportWorkbookOptionalHeaderCellStyler = PxlConstants.DEFAULT_EXPORT_WORKBOOK_OPTIONAL_HEADER_CELL_STYLER;
        }

        Class<? extends PxlStyler> exportWorkbookDataCellStyler = null;
        if (Objects.nonNull(workbookOption) && Objects.nonNull(workbookOption.getExportWorkbookDataCellStyler())) {
            exportWorkbookDataCellStyler = workbookOption.getExportWorkbookDataCellStyler();
        } else if (Objects.nonNull(workbookAnnotation) && Objects.nonNull(workbookAnnotation.exportWorkbookDataCellStyler())) {
            exportWorkbookDataCellStyler = workbookAnnotation.exportWorkbookDataCellStyler();
        }
        if (!PxlMiscUtils.isEffectiveCellStylerClass(exportWorkbookDataCellStyler)) {
            exportWorkbookDataCellStyler = PxlConstants.DEFAULT_EXPORT_WORKBOOK_DATA_CELL_STYLER;
        }

        ResourceBundle exportResourceBundle = Optional.ofNullable(workbookOption)
                .map(PxlExportWorkbookOption::getExportResourceBundle)
                .orElse(null);
        if (Objects.isNull(exportResourceBundle) && Objects.nonNull(workbookAnnotation)) {
            exportResourceBundle = PxlI18nContent.loadBundle(workbookAnnotation.exportI18nBaseName(),
                    workbookAnnotation.exportI18nLanguage(), workbookAnnotation.exportI18nCountry());
        }

        final List<PxlExportSheetOption> exportSheetOptions = Optional.ofNullable(workbookOption)
                .map(option -> option.getExportSheetOptions())
                .orElseGet(ArrayList::new);

        final List<PxlExportSheetMeta> exportSheetMetas = new ArrayList<>();

        // A CSV destination writes no POI workbook, so neither it nor its formula evaluator is created.
        final Workbook workbook = forCsv ? null : PxlWorkbookSupport.createWorkbook(exportExcelEngine, exportSXSSFRowAccessWindowSize);

        final FormulaEvaluator formulaEvaluator = forCsv ? null : PxlWorkbookUtils.createFormulaEvaluator(workbook);

        return new PxlExportWorkbookMeta(
                workbook,
                formulaEvaluator,
                exportExcelEngine,
                exportFileFormat,
                exportPassword,
                exportDataValidation,
                exportSXSSFRowAccessWindowSize,
                exportCsvCharset,
                exportCsvDelimiter,
                exportCsvBom,
                exportWorkbookRequiredHeaderCellStyler,
                exportWorkbookOptionalHeaderCellStyler,
                exportWorkbookDataCellStyler,
                exportResourceBundle,
                exportSheetOptions,
                exportSheetMetas
        );
    }

    /**
     * Returns the export sheet option registered at the given index, or {@code null} if absent.
     *
     * @param index the zero-based position of the sheet option
     * @return the sheet option at the index, or {@code null} if out of range
     */
    public PxlExportSheetOption getExportSheetOption(final int index) {

        return PxlCollectionUtils.get(this.exportSheetOptions, index);
    }

    /**
     * Appends the given sheet metadata to this workbook's sheet metadata list.
     *
     * @param exportSheetMetas the sheet metadata to add
     */
    public void addExportSheetMetas(final List<PxlExportSheetMeta> exportSheetMetas) {

        this.exportSheetMetas.addAll(exportSheetMetas);
    }

    /**
     * Appends a single sheet metadata to this workbook's sheet metadata list.
     *
     * @param exportSheetMeta the sheet metadata to add
     */
    public void addExportSheetMeta(final PxlExportSheetMeta exportSheetMeta) {

        this.exportSheetMetas.add(exportSheetMeta);
    }

    /**
     * Returns a cached {@link CellStyle} produced by the given styler class, creating and caching it on first use.
     * Returns {@code null} when the styler class is ineffective, previously failed, or fails to instantiate/apply
     * (the failure is recorded and diagnostics printed once to short-circuit repeated attempts and prevent style bloat).
     *
     * @param cellStylerClass the styler class whose cell style is requested
     * @return the reusable cell style, or {@code null} if none could be produced
     */
    public CellStyle getCellStyle(final Class<? extends PxlStyler> cellStylerClass) {

        if (!PxlMiscUtils.isEffectiveCellStylerClass(cellStylerClass)) {
            return null;
        }

        // Short-circuit immediately for stylers that previously failed to apply. (Prevents workbook style count explosion from retrying createCellStyle per cell)
        if (failedCellStyleSet.contains(cellStylerClass)) {
            return null;
        }

        CellStyle cellStyle = cellStyleMap.get(cellStylerClass);

        if (Objects.isNull(cellStyle)) {
            try {
                final PxlStyler styler = (PxlStyler) PxlReflectionSupport.newClassInstance(cellStylerClass);
                cellStyle = this.workbook.createCellStyle();
                styler.apply(this.workbook, cellStyle);
                cellStyleMap.put(cellStylerClass, cellStyle);
            } catch (Exception e) {
                // Record the failure to short-circuit later calls (prevent bloat), and report the cause once (diagnostic).
                // Log the diagnostic via SLF4J, and continue exporting that cell without a style.
                failedCellStyleSet.add(cellStylerClass);
                LOGGER.warn(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.LOG_STYLER_APPLY_FAILED, cellStylerClass.getSimpleName(), e.getMessage()));
                return null;
            }
        }

        return cellStyle;
    }

    /**
     * Sets the given string value on the cell using a quote-prefixed cell style, so numeric-looking text stays text.
     * The quote-prefixed style is derived from the cell's current base style and cached per base style index.
     *
     * @param cell  the target cell; returned unchanged when {@code null}
     * @param value the string value to write
     * @return the same cell
     */
    public Cell setQuotePrefixedCellValue(final Cell cell, final String value) {

        if (Objects.isNull(cell)) {
            return cell;
        }

        final CellStyle quotePrefixedCellStyle = getOrCreateQuotePrefixedCellStyle(cell);

        if (Objects.nonNull(quotePrefixedCellStyle)) {
            cell.setCellStyle(quotePrefixedCellStyle);
        }

        cell.setCellValue(value);

        return cell;
    }

    // Obtain from the given (per-workbook) cache, or create only once, a style that applies quote-prefix to the base style (the cell style at call time).
    // For data cells the exporter first applies a per-column shared data style to the cell, so the base style index is stable across the whole column.

    /**
     * Obtains from the per-workbook cache, or creates once, a style that clones the cell's current base style and
     * enables the quote prefix. A {@code null} result (creation failure) is also cached to avoid retry-driven style bloat.
     *
     * @param cell the cell whose base style is cloned; {@code null} yields {@code null}
     * @return the quote-prefixed cell style, or {@code null} if it could not be created
     */
    private CellStyle getOrCreateQuotePrefixedCellStyle(final Cell cell) {

        if (Objects.isNull(cell)) {
            return null;
        }

        final Workbook workbook = cell.getSheet().getWorkbook();
        final CellStyle baseCellStyle = cell.getCellStyle();
        final int baseCellStyleIndex = baseCellStyle.getIndex();

        // Cache creation failures (null) too, to prevent retries on later cells (-> style bloat).
        if (quotePrefixedStyleCache.containsKey(baseCellStyleIndex)) {
            return quotePrefixedStyleCache.get(baseCellStyleIndex);
        }

        CellStyle quotePrefixedCellStyle = null;

        try {
            quotePrefixedCellStyle = workbook.createCellStyle();
            quotePrefixedCellStyle.cloneStyleFrom(baseCellStyle);
            quotePrefixedCellStyle.setQuotePrefixed(true);
        } catch (IllegalStateException e) {
            // Style creation failed, e.g. due to exceeding the workbook cell-style count limit. Log the diagnostic once,
            // cache the failure to prevent bloat, and continue exporting those cells without quote-prefix.
            LOGGER.warn(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.LOG_QUOTE_PREFIX_STYLE_FAILED, e.getMessage()));
        }

        quotePrefixedStyleCache.put(baseCellStyleIndex, quotePrefixedCellStyle);

        return quotePrefixedCellStyle;
    }

    // Obtain from the given (per-workbook) cache, or create only once per combination, a style that inherits the (column-shared)
    // data style already applied to the cell while adding only a date display format.
    // For data cells the exporter first applies a per-column shared data style, so the base style index is stable across the whole column.

    /**
     * Returns a cell style that inherits the cell's current (column-shared) base style and adds the given Excel date
     * display format, obtained from or created once in the per-workbook cache keyed by base style index and format code.
     * Creation failures (e.g. exceeding the workbook cell-style limit) are cached as {@code null} to prevent style bloat.
     *
     * @param cell            the cell whose base style is inherited; {@code null} yields {@code null}
     * @param excelFormatCode the Excel number/date format code to apply; blank yields {@code null}
     * @return the date-formatted cell style, or {@code null} if it could not be created
     */
    public CellStyle getOrCreateDateFormattedCellStyle(final Cell cell,
                                                       final String excelFormatCode) {

        if (Objects.isNull(cell) || StringUtils.isBlank(excelFormatCode)) {
            return null;
        }

        final Workbook workbook = cell.getSheet().getWorkbook();
        final CellStyle baseCellStyle = cell.getCellStyle();
        final int baseCellStyleIndex = baseCellStyle.getIndex();
        final String cacheKey = baseCellStyleIndex + "|" + excelFormatCode;

        // Cache creation failures (null) too, to prevent retries on later cells (-> style bloat).
        if (dateFormattedStyleCache.containsKey(cacheKey)) {
            return dateFormattedStyleCache.get(cacheKey);
        }

        CellStyle dateFormattedCellStyle = null;

        try {
            dateFormattedCellStyle = workbook.createCellStyle();
            dateFormattedCellStyle.cloneStyleFrom(baseCellStyle);
            final short dataFormat = workbook.getCreationHelper().createDataFormat().getFormat(excelFormatCode);
            dateFormattedCellStyle.setDataFormat(dataFormat);
        } catch (IllegalStateException e) {
            // Style creation failed, e.g. due to exceeding the workbook cell-style count limit. Log the diagnostic once,
            // cache the failure to prevent bloat, and continue exporting those cells as Numeric without a display format (General).
            LOGGER.warn(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.LOG_DATE_NUMERIC_STYLE_FAILED, e.getMessage()));
        }

        dateFormattedStyleCache.put(cacheKey, dateFormattedCellStyle);

        return dateFormattedCellStyle;
    }

    /**
     * Returns whether any sheet in this workbook has at least one column exported as a formula.
     *
     * @return {@code true} if any sheet has an export-string-as-formula column
     */
    public boolean hasAnyExportStringAsFormulaColumn() {

        return PxlCollectionUtils.emptyIfNull(this.exportSheetMetas)
                .stream()
                .anyMatch(PxlExportSheetMeta::hasAnyExportStringAsFormulaColumn);
    }

}
