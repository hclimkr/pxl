package io.github.hclimkr.pxl.option;

import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import lombok.*;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.ResourceBundle;

/**
 * Excel workbook import option
 */
@Getter
@Setter
@NoArgsConstructor
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
     * Specifies whether to use the stream reader on import. (Works only with XSSF-format Excel files.)
     * https://github.com/pjfanning/excel-streaming-reader
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
     * Specifies the character encoding set of the CSV to be imported.
     * https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html
     */
    @Builder.Default
    private final String importCsvCharset = null;

    /**
     * Specifies the delimiter of the CSV to be imported.
     */
    @Builder.Default
    private final Character importCsvDelimiter = null;

    /**
     * Specifies the resource bundle for multilingual support on import.
     */
    @Builder.Default
    private ResourceBundle importResourceBundle = null;

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
     * @return {@code true} (as specified by {@link java.util.Collection#add})
     * @throws PxlNullPointerException if {@code importSheetOption} is {@code null}
     */
    public boolean addImportSheetOption(final PxlImportSheetOption importSheetOption)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(importSheetOption, "importSheetOption");

        return this.importSheetOptions.add(importSheetOption);
    }

}
