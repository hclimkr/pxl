package io.github.hclimkr.pxl.option;

import io.github.hclimkr.pxl.PxlExcelEngine;
import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.exception.PxlI18nException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.styler.PxlStyler;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;

import java.util.*;

/**
 * Excel workbook export option
 */
@Getter
@AllArgsConstructor
@Builder
public final class PxlExportWorkbookOption {

    /**
     * Specifies the POI engine used to write the Excel workbook on export.
     */
    @Builder.Default
    private final PxlExcelEngine exportExcelEngine = null;

    /**
     * Specifies the password used to protect the document on export.
     */
    @Builder.Default
    private final String exportPassword = null;

    /**
     * Specifies whether to bean-validate the data to be exported.
     * <p>
     * This is Bean Validation over the objects being written, not Excel's "data validation" feature — the
     * dropdown written into the file comes from {@link PxlColumn#exportOptionItems()} instead.
     */
    @Builder.Default
    private final Boolean exportDataValidation = null;

    /**
     * Specifies the rowAccessWindowSize value used when exporting with SXSSF.
     */
    @Builder.Default
    private final Integer exportSXSSFRowAccessWindowSize = null;

    /**
     * Specifies the style applied to required header cells on export.
     */
    @Builder.Default
    private final Class<? extends PxlStyler> exportWorkbookRequiredHeaderCellStyler = null;

    /**
     * Specifies the style applied to optional header cells on export.
     */
    @Builder.Default
    private final Class<? extends PxlStyler> exportWorkbookOptionalHeaderCellStyler = null;

    /**
     * Specifies the style applied to data cells on export.
     */
    @Builder.Default
    private final Class<? extends PxlStyler> exportWorkbookDataCellStyler = null;

    /**
     * Specifies the resource bundle for multilingual support on export.
     * <p>
     * Takes precedence over {@link PxlWorkbook#exportI18nBaseName()} and its language/country pair: given a bundle
     * here, PXL uses it and never loads the annotated one, so a base name that resolves to nothing cannot raise
     * {@link PxlI18nException}. Use it when the bundle comes from somewhere an annotation cannot name — one the
     * application already resolved for the current request, say. Left {@code null}, the annotation decides.
     */
    @Builder.Default
    private final ResourceBundle exportResourceBundle = null;

    /**
     * Per-sheet export overrides, matched to sheet fields by name. Empty by default.
     */
    @Builder.Default
    private final List<PxlExportSheetOption> exportSheetOptions = new ArrayList<>();

    /**
     * Returns the sheet export option at the given position in this workbook option.
     *
     * @param index the zero-based position within the sheet option list
     * @return the sheet option at that position, or {@code null} if the index is out of range
     */
    public PxlExportSheetOption getExportSheetOption(final int index) {

        return PxlCollectionUtils.get(this.exportSheetOptions, index);
    }

    /**
     * Null-safe accessor that returns the sheet export option at the given position from the supplied workbook option.
     *
     * @param exportWorkbookOption the workbook option to read from; may be {@code null}
     * @param index                the zero-based position within the sheet option list
     * @return the sheet option at that position, or {@code null} if the workbook option is {@code null} or the index is out of range
     */
    public static PxlExportSheetOption getExportSheetOption(@Nullable final PxlExportWorkbookOption exportWorkbookOption,
                                                            final int index) {

        return Optional.ofNullable(exportWorkbookOption)
                .map(workbookOption -> workbookOption.getExportSheetOption(index))
                .orElse(null);
    }

    /**
     * Null-safe accessor that returns the sheet export options of the supplied workbook option.
     *
     * @param exportWorkbookOption the workbook option to read from; may be {@code null}
     * @return the list of sheet options, or an empty list if the workbook option is {@code null}
     */
    public static List<PxlExportSheetOption> getExportSheetOptions(@Nullable final PxlExportWorkbookOption exportWorkbookOption) {

        return Optional.ofNullable(exportWorkbookOption)
                .map(workbookOption -> workbookOption.getExportSheetOptions())
                .orElseGet(ArrayList::new);
    }

    /**
     * Appends a sheet export option to this workbook option.
     *
     * @param exportSheetOption the sheet option to append
     * @return {@code true} (as specified by {@link Collection#add})
     * @throws PxlNullPointerException if {@code exportSheetOption} is {@code null}
     */
    public boolean addExportSheetOption(final PxlExportSheetOption exportSheetOption)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(exportSheetOption, "exportSheetOption");

        return this.exportSheetOptions.add(exportSheetOption);
    }

}
