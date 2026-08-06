package io.github.hclimkr.pxl.option;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.exception.PxlI18nException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;

import java.util.*;

/**
 * Workbook import option, for an Excel or a CSV source alike — the attributes an Excel import has no use for
 * (the CSV charset and delimiter) are ignored there rather than belonging to a separate option type.
 */
@Getter
@AllArgsConstructor
@Builder
public final class PxlImportWorkbookOption {

    /**
     * Specifies the password used to remove document protection on import.
     */
    @Builder.Default
    private final String importPassword = null;

    /**
     * Specifies whether to validate the data to be imported.
     */
    @Builder.Default
    private final Boolean importDataValidation = null;

    /**
     * Specifies whether to use the stream reader on import. (Works only with XLSX files.)
     *
     * @see <a href="https://github.com/pjfanning/excel-streaming-reader">excel-streaming-reader</a>
     */
    @Builder.Default
    private final Boolean importUsingStreamReader = null;

    /**
     * Specifies the RowCacheSize value when importing with the stream reader.
     */
    @Builder.Default
    private final Integer importStreamReaderRowCacheSize = null;

    /**
     * Specifies the BufferSize value when importing with the stream reader.
     */
    @Builder.Default
    private final Integer importStreamReaderBufferSize = null;

    /**
     * Specifies the character encoding set of the CSV to be imported, for every sheet of the workbook.
     * Ignored for an Excel source, which is one file whose encoding the format itself carries.
     * <p>
     * Overrides {@link PxlWorkbook#importCsvCharset()} but not {@link PxlSheet#importCsvCharset()}, since a CSV
     * workbook is read as one file per sheet. Left {@code null} or blank, the workbook says nothing and the levels
     * below decide, ending at {@link PxlConstants#DEFAULT_IMPORT_CSV_CHARSET} ({@code "UTF-8"}).
     * The per-sheet counterpart is {@code PxlImportSheetOption.importCsvCharset}.
     *
     * @see <a href="https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html">Java supported encodings</a>
     */
    @Builder.Default
    private final String importCsvCharset = null;

    /**
     * Specifies the delimiter of the CSV to be imported, for every sheet of the workbook.
     * Ignored for an Excel source, which has no delimiter.
     * <p>
     * Overrides {@link PxlWorkbook#importCsvDelimiter()} but not {@link PxlSheet#importCsvDelimiter()}, since a CSV
     * workbook is read as one file per sheet. Left {@code null} or NUL, the workbook says nothing and the levels
     * below decide, ending at {@link PxlConstants#DEFAULT_IMPORT_CSV_DELIMITER} ({@code ','}).
     * The per-sheet counterpart is {@code PxlImportSheetOption.importCsvDelimiter}.
     */
    @Builder.Default
    private final Character importCsvDelimiter = null;

    /**
     * Specifies the resource bundle for multilingual support on import.
     * <p>
     * Takes precedence over {@link PxlWorkbook#importI18nBaseName()} and its language/country pair: given a bundle
     * here, PXL uses it and never loads the annotated one, so a base name that resolves to nothing cannot raise
     * {@link PxlI18nException}. Use it when the bundle comes from somewhere an annotation cannot name — one the
     * application already resolved for the current request, say. Left {@code null}, the annotation decides.
     */
    @Builder.Default
    private final ResourceBundle importResourceBundle = null;

    /**
     * Per-sheet import overrides, matched to sheet fields by name. Empty by default.
     */
    @Builder.Default
    private final List<PxlImportSheetOption> importSheetOptions = new ArrayList<>();

    /**
     * Returns the sheet import option at the given position in this workbook option.
     *
     * @param index the zero-based position within the sheet option list
     * @return the sheet option at that position, or {@code null} if the index is out of range
     */
    public PxlImportSheetOption getImportSheetOption(final int index) {

        return PxlCollectionUtils.get(this.importSheetOptions, index);
    }

    /**
     * Null-safe accessor that returns the sheet import option at the given position from the supplied workbook option.
     *
     * @param importWorkbookOption the workbook option to read from; may be {@code null}
     * @param index                the zero-based position within the sheet option list
     * @return the sheet option at that position, or {@code null} if the workbook option is {@code null} or the index is out of range
     */
    public static PxlImportSheetOption getImportSheetOption(@Nullable final PxlImportWorkbookOption importWorkbookOption,
                                                            final int index) {

        return Optional.ofNullable(importWorkbookOption)
                .map(workbookOption -> workbookOption.getImportSheetOption(index))
                .orElse(null);
    }

    /**
     * Null-safe accessor that returns the sheet import options of the supplied workbook option.
     *
     * @param importWorkbookOption the workbook option to read from; may be {@code null}
     * @return the list of sheet options, or an empty list if the workbook option is {@code null}
     */
    public static List<PxlImportSheetOption> getImportSheetOptions(@Nullable final PxlImportWorkbookOption importWorkbookOption) {

        return Optional.ofNullable(importWorkbookOption)
                .map(workbookOption -> workbookOption.getImportSheetOptions())
                .orElseGet(ArrayList::new);
    }

    /**
     * Appends a sheet import option to this workbook option.
     *
     * @param importSheetOption the sheet option to append
     * @return {@code true} (as specified by {@link Collection#add})
     * @throws PxlNullPointerException if {@code importSheetOption} is {@code null}
     */
    public boolean addImportSheetOption(final PxlImportSheetOption importSheetOption)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(importSheetOption, "importSheetOption");

        return this.importSheetOptions.add(importSheetOption);
    }

}
