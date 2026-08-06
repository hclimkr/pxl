package io.github.hclimkr.pxl.internal.meta;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.PxlFileFormat;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.exception.PxlDataException;
import io.github.hclimkr.pxl.exception.PxlI18nException;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nContent;
import io.github.hclimkr.pxl.internal.support.PxlWorkbookSupport;
import io.github.hclimkr.pxl.option.PxlImportSheetOption;
import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import lombok.Getter;
import lombok.Setter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.DataFormatter;

import java.util.*;

/**
 * Excel workbook import metadata.
 */
@Getter
public final class PxlImportWorkbookMeta {

    @Setter
    private PxlFileFormat importFileFormat;    // Not supplied via annotation or option; it is detected.

    private final String importPassword;

    private final boolean importDataValidation;

    private final boolean importUsingStreamReader;

    private final int importStreamReaderRowCacheSize;

    private final int importStreamReaderBufferSize;

    private final String importCsvCharset;

    private final char importCsvDelimiter;

    private final ResourceBundle importResourceBundle;

    private final List<PxlImportSheetOption> importSheetOptions;

    private final List<PxlImportSheetMeta> importSheetMetas;

    private final DataFormatter importDataFormatterCache;

    /**
     * Creates the resolved import metadata for a workbook, storing the merged option/annotation values and a
     * fresh POI {@link DataFormatter} cache. The file format is detected later from the actual source.
     *
     * @param importPassword                 the password for encrypted workbooks; may be {@code null}
     * @param importDataValidation           whether bean validation is applied
     * @param importUsingStreamReader        whether the streaming reader is used
     * @param importStreamReaderRowCacheSize the streaming reader row cache size
     * @param importStreamReaderBufferSize   the streaming reader buffer size
     * @param importCsvCharset               the CSV charset name
     * @param importCsvDelimiter             the CSV delimiter character
     * @param importResourceBundle           the content i18n bundle for sheet/column name translation; may be {@code null}
     * @param importSheetOptions             the per-sheet import overrides
     * @param importSheetMetas               the (initially empty) sheet metadata list
     */
    private PxlImportWorkbookMeta(final String importPassword,
                                  final boolean importDataValidation,
                                  final boolean importUsingStreamReader,
                                  final int importStreamReaderRowCacheSize,
                                  final int importStreamReaderBufferSize,
                                  final String importCsvCharset,
                                  final char importCsvDelimiter,
                                  final ResourceBundle importResourceBundle,
                                  final List<PxlImportSheetOption> importSheetOptions,
                                  final List<PxlImportSheetMeta> importSheetMetas) {

        this.importPassword = importPassword;
        this.importDataValidation = importDataValidation;
        this.importUsingStreamReader = importUsingStreamReader;
        this.importStreamReaderRowCacheSize = importStreamReaderRowCacheSize;
        this.importStreamReaderBufferSize = importStreamReaderBufferSize;
        this.importCsvCharset = importCsvCharset;
        this.importCsvDelimiter = importCsvDelimiter;
        this.importResourceBundle = importResourceBundle;
        this.importSheetOptions = importSheetOptions;
        this.importSheetMetas = importSheetMetas;
        // Locale.ROOT keeps numeric-cell -> String rendering deterministic (locale-independent decimal/grouping symbols).
        this.importDataFormatterCache = new DataFormatter(Locale.ROOT);
    }

    /**
     * On import, collects the workbook metadata from the workbook option and the workbook class.
     * The workbook option takes precedence over the workbook class.
     * The file format is not resolved here; it is detected later from the actual source.
     *
     * @param workbookClass  the {@link PxlWorkbook}-annotated workbook class supplying annotation defaults; may be {@code null}
     * @param workbookOption runtime overrides taking precedence over the class annotation; may be {@code null}
     * @return the assembled import workbook metadata, holding resolved password/validation/stream-reader/CSV settings, i18n bundle and sheet options
     * @throws PxlDataException if the workbook name field type is invalid
     * @throws PxlI18nException if the import content i18n bundle cannot be found for the configured base name and locale
     */
    public static PxlImportWorkbookMeta makeImportWorkbookMeta(@Nullable final Class<?> workbookClass,
                                                               @Nullable final PxlImportWorkbookOption workbookOption)
            throws PxlDataException, PxlI18nException {

        PxlWorkbookSupport.validateWorkbookNameFieldType(workbookClass);

        final PxlWorkbook workbookAnnotation = Optional.ofNullable(workbookClass)
                .map(c -> c.getAnnotation(PxlWorkbook.class))
                .orElse(null);

        final String importPassword = Optional.ofNullable(workbookOption)
                .flatMap(option -> Optional.ofNullable(option.getImportPassword()))
                .orElseGet(() -> Optional.ofNullable(workbookAnnotation)
                        .map(PxlWorkbook::importPassword)
                        .orElse(PxlConstants.DEFAULT_IMPORT_PASSWORD));

        final boolean importDataValidation = Optional.ofNullable(workbookOption)
                .flatMap(option -> Optional.ofNullable(option.getImportDataValidation()))
                .orElseGet(() -> Optional.ofNullable(workbookAnnotation)
                        .map(PxlWorkbook::importDataValidation)
                        .orElse(PxlConstants.DEFAULT_IMPORT_DATA_VALIDATION));

        final boolean importUsingStreamReader = Optional.ofNullable(workbookOption)
                .flatMap(option -> Optional.ofNullable(option.getImportUsingStreamReader()))
                .orElseGet(() -> Optional.ofNullable(workbookAnnotation)
                        .map(PxlWorkbook::importUsingStreamReader)
                        .orElse(PxlConstants.DEFAULT_IMPORT_USING_STREAM_READER));

        final int importStreamReaderRowCacheSize = Optional.ofNullable(workbookOption)
                .flatMap(option -> Optional.ofNullable(option.getImportStreamReaderRowCacheSize()))
                .orElseGet(() -> Optional.ofNullable(workbookAnnotation)
                        .map(PxlWorkbook::importStreamReaderRowCacheSize)
                        .orElse(PxlConstants.DEFAULT_IMPORT_STREAM_READER_ROW_CACHE_SIZE));

        final int importStreamReaderBufferSize = Optional.ofNullable(workbookOption)
                .flatMap(option -> Optional.ofNullable(option.getImportStreamReaderBufferSize()))
                .orElseGet(() -> Optional.ofNullable(workbookAnnotation)
                        .map(PxlWorkbook::importStreamReaderBufferSize)
                        .orElse(PxlConstants.DEFAULT_IMPORT_STREAM_READER_BUFFER_SIZE));

        final String importCsvCharset = Optional.ofNullable(workbookOption)
                .flatMap(option -> Optional.ofNullable(option.getImportCsvCharset()))
                .filter(StringUtils::isNotBlank)
                .orElseGet(() -> Optional.ofNullable(workbookAnnotation)
                        .map(PxlWorkbook::importCsvCharset)
                        .filter(StringUtils::isNotBlank)
                        .orElse(PxlConstants.DEFAULT_IMPORT_CSV_CHARSET));

        final char importCsvDelimiter = Optional.ofNullable(workbookOption)
                .flatMap(option -> Optional.ofNullable(option.getImportCsvDelimiter()))
                .filter(delimiter -> delimiter != PxlConstants.UNSPECIFIED_IMPORT_CSV_DELIMITER)
                .orElseGet(() -> Optional.ofNullable(workbookAnnotation)
                        .map(PxlWorkbook::importCsvDelimiter)
                        .filter(delimiter -> delimiter != PxlConstants.UNSPECIFIED_IMPORT_CSV_DELIMITER)
                        .orElse(PxlConstants.DEFAULT_IMPORT_CSV_DELIMITER));

        ResourceBundle importResourceBundle = Optional.ofNullable(workbookOption)
                .map(PxlImportWorkbookOption::getImportResourceBundle)
                .orElse(null);
        if (Objects.isNull(importResourceBundle) && Objects.nonNull(workbookAnnotation)) {
            importResourceBundle = PxlI18nContent.loadBundle(workbookAnnotation.importI18nBaseName(),
                    workbookAnnotation.importI18nLanguage(), workbookAnnotation.importI18nCountry());
        }

        final List<PxlImportSheetOption> importSheetOptions = Optional.ofNullable(workbookOption)
                .map(option -> option.getImportSheetOptions())
                .orElseGet(ArrayList::new);

        final List<PxlImportSheetMeta> importSheetMetas = new ArrayList<>();

        return new PxlImportWorkbookMeta(
                importPassword,
                importDataValidation,
                importUsingStreamReader,
                importStreamReaderRowCacheSize,
                importStreamReaderBufferSize,
                importCsvCharset,
                importCsvDelimiter,
                importResourceBundle,
                importSheetOptions,
                importSheetMetas
        );
    }

    /**
     * Returns the import sheet option registered at the given index, or {@code null} if absent.
     *
     * @param index the zero-based position of the sheet option
     * @return the sheet option at the index, or {@code null} if out of range
     */
    public PxlImportSheetOption getImportSheetOption(final int index) {

        return PxlCollectionUtils.get(this.importSheetOptions, index);
    }

    /**
     * Appends the given sheet metadata to this workbook's sheet metadata list.
     *
     * @param importSheetMetas the sheet metadata to add
     */
    public void addImportSheetMetas(final List<PxlImportSheetMeta> importSheetMetas) {

        this.importSheetMetas.addAll(importSheetMetas);
    }

    /**
     * Appends a single sheet metadata to this workbook's sheet metadata list.
     *
     * @param importSheetMeta the sheet metadata to add
     */
    public void addImportSheetMeta(final PxlImportSheetMeta importSheetMeta) {

        this.importSheetMetas.add(importSheetMeta);
    }

}
