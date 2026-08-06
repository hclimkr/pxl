package io.github.hclimkr.pxl.option;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.styler.PxlStyler;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NonNull;
import org.apache.commons.lang3.StringUtils;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

/**
 * Excel sheet export option
 */
@Getter
@AllArgsConstructor
@Builder
public final class PxlExportSheetOption {

    /**
     * The field name for the sheet object
     */
    @NonNull
    @Builder.Default
    private final String fieldName = PxlConstants.SHEET_FIELD_NAME_WILD_CARD;

    /**
     * Specifies the name of the sheet on export.
     * <p>
     * Doubles as a content-i18n key, just like the {@link PxlSheet#name()} it overrides:
     * when the workbook sets {@code exportI18nBaseName}, the name is resolved through that bundle first and it is the
     * translation that is written. A name the bundle does not carry is used as it stands.
     */
    @Builder.Default
    private final List<String> exportSheetNames = null;

    /**
     * Specifies whether to export.
     */
    @Builder.Default
    private final Boolean exportEnabled = null;

    /**
     * Specifies whether to export as a sample.
     */
    @Builder.Default
    private final Boolean exportSampleEnabled = null;

    /**
     * Specifies whether to override a superclass field that uses the same sheet name on export, if one exists.
     */
    @Builder.Default
    private final Boolean exportOverrideSuperClassSheet = null;

    /**
     * Specifies the height of rows within the sheet on export.
     */
    @Builder.Default
    private final Float exportRowHeightInPoints = null;

    /**
     * Specifies the ordering among sheets on export. (in alphabetical order)
     */
    @Builder.Default
    private final String exportOrder = null;

    /**
     * Specifies the name of the field by which to group and split into multiple sheets on export.
     */
    @Builder.Default
    private final String exportGroupingFieldName = null;

    /**
     * Specifies the index of the row used as the header on export.
     * (The default is the first row. When set explicitly, use a 1-based value, and it must be less than the value of exportFirstDataRowIndex.)
     */
    @Builder.Default
    private final Integer exportHeaderRowIndex = null;

    /**
     * Specifies the index of the starting row used as data on export.
     * (The default is the second row. When set explicitly, use a 1-based value, and it must be greater than the value of exportHeaderRowIndex and less than or equal to the value of exportLastDataRowIndex.)
     */
    @Builder.Default
    private final Integer exportFirstDataRowIndex = null;

    /**
     * Specifies the index of the ending row used as data on export.
     * (The default is the last row. When set explicitly, use a 1-based value, and it must be greater than or equal to the value of exportFirstDataRowIndex.)
     */
    @Builder.Default
    private final Integer exportLastDataRowIndex = null;

    /**
     * Specifies the index of the starting column used as data on export.
     * (The default is the first column. When set explicitly, use a 1-based value, and it must be less than or equal to the value of exportLastDataColumnIndex.)
     */
    @Builder.Default
    private final Integer exportFirstDataColumnIndex = null;

    /**
     * Specifies the index of the ending column used as data on export.
     * (The default is the last column. When set explicitly, use a 1-based value, and it must be greater than or equal to the value of exportFirstDataColumnIndex.)
     */
    @Builder.Default
    private final Integer exportLastDataColumnIndex = null;

    /**
     * Specifies whether to export when the data list is null.
     */
    @Builder.Default
    private final Boolean exportIfNull = null;

    /**
     * Specifies whether to export when the data list is empty.
     */
    @Builder.Default
    private final Boolean exportIfEmpty = null;

    /**
     * Specifies whether to apply a filter on export.
     */
    @Builder.Default
    private final Boolean exportColumnFilter = null;

    /**
     * Specifies the style applied to required header cells on export.
     */
    @Builder.Default
    private final Class<? extends PxlStyler> exportSheetRequiredHeaderCellStyler = null;

    /**
     * Specifies the style applied to optional header cells on export.
     */
    @Builder.Default
    private final Class<? extends PxlStyler> exportSheetOptionalHeaderCellStyler = null;

    /**
     * Specifies the style applied to data cells on export.
     */
    @Builder.Default
    private final Class<? extends PxlStyler> exportSheetDataCellStyler = null;

    /**
     * Per-column export overrides, matched to column fields by name. Empty by default.
     */
    @Builder.Default
    private final List<PxlExportColumnOption> exportColumnOptions = new ArrayList<>();

    /**
     * Null-safe accessor that returns the export sheet names of the supplied sheet option, trimmed with blank entries removed.
     *
     * @param sheetOption the sheet option to read from; may be {@code null}
     * @return the trimmed, non-blank sheet names, or {@code null} if the sheet option or its names are {@code null}
     */
    public static List<String> getExportSheetNames(@Nullable final PxlExportSheetOption sheetOption) {

        return Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getExportSheetNames())
                        .map(names -> names.stream()
                                .map(StringUtils::trim)
                                .filter(StringUtils::isNotBlank)
                                .collect(Collectors.toList())))
                .orElse(null);
    }

    /**
     * Returns the column export option at the given position in this sheet option.
     *
     * @param index the zero-based position within the column option list
     * @return the column option at that position, or {@code null} if the index is out of range
     */
    public PxlExportColumnOption getExportColumnOption(final int index) {

        return PxlCollectionUtils.get(this.exportColumnOptions, index);
    }

    /**
     * Null-safe accessor that returns the column export option at the given position from the supplied sheet option.
     *
     * @param exportSheetOption the sheet option to read from; may be {@code null}
     * @param index             the zero-based position within the column option list
     * @return the column option at that position, or {@code null} if the sheet option is {@code null} or the index is out of range
     */
    public static PxlExportColumnOption getExportColumnOption(@Nullable final PxlExportSheetOption exportSheetOption,
                                                              final int index) {

        return Optional.ofNullable(exportSheetOption)
                .map(sheetOption -> sheetOption.getExportColumnOption(index))
                .orElse(null);
    }

    /**
     * Null-safe accessor that returns the column export options of the supplied sheet option.
     *
     * @param exportSheetOption the sheet option to read from; may be {@code null}
     * @return the list of column options, or an empty list if the sheet option is {@code null}
     */
    public static List<PxlExportColumnOption> getExportColumnOptions(@Nullable final PxlExportSheetOption exportSheetOption) {

        return Optional.ofNullable(exportSheetOption)
                .map(sheetOption -> sheetOption.getExportColumnOptions())
                .orElseGet(ArrayList::new);
    }

    /**
     * Appends a column export option to this sheet option.
     *
     * @param exportColumnOption the column option to append
     * @return {@code true} (as specified by {@link Collection#add})
     * @throws PxlNullPointerException if {@code exportColumnOption} is {@code null}
     */
    public boolean addExportColumnOption(final PxlExportColumnOption exportColumnOption)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(exportColumnOption, "exportColumnOption");

        return this.exportColumnOptions.add(exportColumnOption);
    }

}
