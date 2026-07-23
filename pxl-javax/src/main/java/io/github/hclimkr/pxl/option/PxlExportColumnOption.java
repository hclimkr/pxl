package io.github.hclimkr.pxl.option;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.styler.PxlStyler;
import lombok.*;
import org.apache.commons.lang3.StringUtils;

import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

/**
 * Excel column export option
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public final class PxlExportColumnOption {

    /**
     * The field name for the column object
     */
    @NonNull    // required
    private String fieldName;

    /**
     * Specifies the name of the column on export.
     */
    @Builder.Default
    private List<String> exportColumnNames = null;

    /**
     * Specifies whether to export.
     */
    @Builder.Default
    private Boolean exportEnabled = null;

    /**
     * Specifies whether to export as a sample.
     */
    @Builder.Default
    private Boolean exportSampleEnabled = null;

    /**
     * Specifies the sample value on export.
     */
    @Builder.Default
    private String exportSample = null;

    /**
     * Specifies whether to trim strings on export.
     */
    @Builder.Default
    private Boolean exportTrim = null;

    /**
     * Specifies the cell formatting string on export.
     * Valid only for fields of type Numeric, Date, LocalTime, LocalDate, LocalDateTime, ZonedDateTime, OffsetTime, OffsetDateTime, and Duration.
     */
    @Builder.Default
    private String exportPattern = null;

    /**
     * Specifies the width of the column on export.
     * in units of 1/256th of a character width (maximum: 255 * 256)
     */
    @Builder.Default
    private Integer exportColumnWidth = null;

    /**
     * Specifies the separator between elements when exporting a collection.
     */
    @Builder.Default
    private String exportCollectionSeparator = null;

    /**
     * Specifies whether to override a superclass field that uses the same column name on export, if one exists.
     */
    @Builder.Default
    private Boolean exportOverrideSuperClassColumn = null;

    /**
     * Specifies the ordering among columns on export. (in alphabetical order)
     */
    @Builder.Default
    private String exportOrder = null;

    /**
     * Specifies the masking rule as a regular expression on export.
     */
    @Builder.Default
    private String exportMasking = null;

    /**
     * Sets the list of selectable options on export.
     */
    @Builder.Default
    private String[] exportOptionItems = null;

    /**
     * Sets an enum field as a drop-down list on export.
     */
    @Builder.Default
    private PxlColumn.ExportEnumDropDownListStyle exportEnumDropDownListStyle = null;

    /**
     * Specifies the string that represents null on export.
     */
    @Builder.Default
    private String exportNullString = null;

    /**
     * Specifies the string that represents the boolean value true on export.
     */
    @Builder.Default
    private String exportTrueString = null;

    /**
     * Specifies the string that represents the boolean value false on export.
     */
    @Builder.Default
    private String exportFalseString = null;

    /**
     * On export, interprets the string as a path to an image and applies the image itself to the cell.
     */
    @Builder.Default
    private Boolean exportStringAsPicture = null;

    /**
     * On export, interprets the string as a formula and applies its computed result itself to the cell.
     */
    @Builder.Default
    private Boolean exportStringAsFormula = null;

    /**
     * Specifies the style applied to required header cells on export.
     */
    @Builder.Default
    private Class<? extends PxlStyler> exportColumnRequiredHeaderCellStyler = null;

    /**
     * Specifies the style applied to optional header cells on export.
     */
    @Builder.Default
    private Class<? extends PxlStyler> exportColumnOptionalHeaderCellStyler = null;

    /**
     * Specifies the style applied to data cells on export.
     */
    @Builder.Default
    private Class<? extends PxlStyler> exportColumnDataCellStyler = null;

    /**
     * Null-safe accessor that returns the export column names of the supplied column option, trimmed with blank entries removed.
     *
     * @param columnOption the column option to read from; may be {@code null}
     * @return the trimmed, non-blank column names, or {@code null} if the column option or its names are {@code null}
     */
    public static List<String> getExportColumnNames(@Nullable final PxlExportColumnOption columnOption) {

        return Optional.ofNullable(columnOption)
                .flatMap(option -> Optional.ofNullable(option.getExportColumnNames())
                        .map(names -> names.stream()
                                .map(StringUtils::trim)
                                .filter(StringUtils::isNotBlank)
                                .collect(Collectors.toList())))
                .orElse(null);
    }

}
