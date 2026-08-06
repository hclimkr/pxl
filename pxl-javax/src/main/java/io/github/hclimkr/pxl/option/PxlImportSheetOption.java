package io.github.hclimkr.pxl.option;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import lombok.*;
import org.apache.commons.lang3.StringUtils;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

/**
 * Sheet import option, for an Excel or a CSV source alike. A CSV workbook is read as one file per sheet, so this
 * is the level at which its charset and delimiter are settled; an Excel import ignores that pair.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public final class PxlImportSheetOption {

    /**
     * The field name for the sheet object
     */
    @NonNull
    @Builder.Default
    private String fieldName = PxlConstants.SHEET_FIELD_NAME_WILD_CARD;

    /**
     * Specifies the name of the sheet on import.
     * <p>
     * Doubles as a content-i18n key, just like the {@link PxlSheet#name()} it overrides:
     * when the workbook sets {@code importI18nBaseName}, the name is resolved through that bundle first and it is the
     * translation that is matched. A name the bundle does not carry is used as it stands.
     */
    @Builder.Default
    private List<String> importSheetNames = null;

    /**
     * Specifies whether to import.
     */
    @Builder.Default
    private Boolean importEnabled = null;

    /**
     * Specifies whether to override a superclass field that uses the same sheet name on import, if one exists.
     */
    @Builder.Default
    private Boolean importOverrideSuperClassSheet = null;

    /**
     * Specifies whether to import hidden rows on import.
     */
    @Builder.Default
    private Boolean importExcludeHiddenRows = null;

    /**
     * Specifies whether to import hidden columns on import.
     */
    @Builder.Default
    private Boolean importExcludeHiddenColumns = null;

    /**
     * Specifies whether, for merged cells, each individual cell is treated as having the same value on import.
     */
    @Builder.Default
    private Boolean importEachCellOfMergedRegion = null;

    /**
     * Specifies the index of the row used as the header on import.
     * (The default is the first row. When set explicitly, use a 1-based value, and it must be less than the value of importFirstDataRowIndex.)
     */
    @Builder.Default
    private Integer importHeaderRowIndex = null;

    /**
     * Specifies the index of the starting row used as data on import.
     * (The default is the second row. When set explicitly, use a 1-based value, and it must be greater than the value of importHeaderRowIndex and less than or equal to the value of importLastDataRowIndex.)
     */
    @Builder.Default
    private Integer importFirstDataRowIndex = null;

    /**
     * Specifies the index of the ending row used as data on import.
     * (The default is the last row. When set explicitly, use a 1-based value, and it must be greater than or equal to the value of importFirstDataRowIndex.)
     */
    @Builder.Default
    private Integer importLastDataRowIndex = null;

    /**
     * Specifies the index of the starting column used as data on import.
     * (The default is the first column. When set explicitly, use a 1-based value, and it must be less than or equal to the value of importLastDataColumnIndex.)
     */
    @Builder.Default
    private Integer importFirstDataColumnIndex = null;

    /**
     * Specifies the index of the ending column used as data on import.
     * (The default is the last column. When set explicitly, use a 1-based value, and it must be greater than or equal to the value of importFirstDataColumnIndex.)
     */
    @Builder.Default
    private Integer importLastDataColumnIndex = null;

    /**
     * Specifies the Character Encoding Set of the CSV to import, for this sheet alone.
     * Ignored for an Excel source, where the whole workbook is one file.
     * <p>
     * Takes precedence over {@link PxlSheet#importCsvCharset()} and, through it, over the workbook charset.
     */
    @Builder.Default
    private String importCsvCharset = null;

    /**
     * Specifies the Delimiter of the CSV to import, for this sheet alone.
     * Ignored for an Excel source, where the whole workbook is one file.
     * <p>
     * Takes precedence over {@link PxlSheet#importCsvDelimiter()} and, through it, over the workbook delimiter.
     */
    @Builder.Default
    private Character importCsvDelimiter = null;

    /**
     * Per-column import overrides, matched to column fields by name. Empty by default.
     */
    @Builder.Default
    private final List<PxlImportColumnOption> importColumnOptions = new ArrayList<>();

    /**
     * Null-safe accessor that returns the import sheet names of the supplied sheet option, with all whitespace removed and blank entries dropped.
     *
     * @param sheetOption the sheet option to read from; may be {@code null}
     * @return the whitespace-stripped, non-blank sheet names, or {@code null} if the sheet option or its names are {@code null}
     */
    public static List<String> getImportSheetNames(@Nullable final PxlImportSheetOption sheetOption) {

        return Optional.ofNullable(sheetOption)
                .flatMap(option -> Optional.ofNullable(option.getImportSheetNames())
                        .map(names -> names.stream()
                                .map(StringUtils::deleteWhitespace)
                                .filter(StringUtils::isNotBlank)
                                .collect(Collectors.toList())))
                .orElse(null);
    }

    /**
     * Returns the column import option at the given position in this sheet option.
     *
     * @param index the zero-based position within the column option list
     * @return the column option at that position, or {@code null} if the index is out of range
     */
    public PxlImportColumnOption getImportColumnOption(final int index) {

        return PxlCollectionUtils.get(this.importColumnOptions, index);
    }

    /**
     * Null-safe accessor that returns the column import option at the given position from the supplied sheet option.
     *
     * @param importSheetOption the sheet option to read from; may be {@code null}
     * @param index             the zero-based position within the column option list
     * @return the column option at that position, or {@code null} if the sheet option is {@code null} or the index is out of range
     */
    public static PxlImportColumnOption getImportColumnOption(@Nullable final PxlImportSheetOption importSheetOption,
                                                              final int index) {

        return Optional.ofNullable(importSheetOption)
                .map(sheetOption -> sheetOption.getImportColumnOption(index))
                .orElse(null);
    }

    /**
     * Null-safe accessor that returns the column import options of the supplied sheet option.
     *
     * @param importSheetOption the sheet option to read from; may be {@code null}
     * @return the list of column options, or an empty list if the sheet option is {@code null}
     */
    public static List<PxlImportColumnOption> getImportColumnOptions(@Nullable final PxlImportSheetOption importSheetOption) {

        return Optional.ofNullable(importSheetOption)
                .map(sheetOption -> sheetOption.getImportColumnOptions())
                .orElseGet(ArrayList::new);
    }

    /**
     * Appends a column import option to this sheet option.
     *
     * @param importColumnOption the column option to append
     * @return {@code true} (as specified by {@link Collection#add})
     * @throws PxlNullPointerException if {@code importColumnOption} is {@code null}
     */
    public boolean addImportColumnOption(final PxlImportColumnOption importColumnOption)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(importColumnOption, "importColumnOption");

        return this.importColumnOptions.add(importColumnOption);
    }

}
