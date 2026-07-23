package io.github.hclimkr.pxl.option;

import io.github.hclimkr.pxl.internal.constraint.Nullable;
import lombok.*;
import org.apache.commons.lang3.StringUtils;

import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

/**
 * Excel column import option
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public final class PxlImportColumnOption {

    /**
     * The field name for the column object
     */
    @NonNull    // required
    private String fieldName;

    /**
     * Specifies the name of the column on import.
     */
    @Builder.Default
    private List<String> importColumnNames = null;

    /**
     * Specifies whether to import.
     */
    @Builder.Default
    private Boolean importEnabled = null;

    /**
     * Specifies whether to trim strings on import.
     */
    @Builder.Default
    private Boolean importTrim = null;

    /**
     * Specifies whether to check the uniqueness of the column's values on import.
     */
    @Builder.Default
    private Boolean importUnique = null;

    /**
     * Specifies the cell formatting string on import.
     * Valid only for fields of type Numeric, Date, LocalTime, LocalDate, LocalDateTime, ZonedDateTime, OffsetTime, and OffsetDateTime.
     */
    @Builder.Default
    private String importPattern = null;

    /**
     * Specifies the string that represents the boolean value true for BOOLEAN cell types on import.
     */
    @Builder.Default
    private String importTrueString = null;

    /**
     * Specifies the string that represents the boolean value false for BOOLEAN cell types on import.
     */
    @Builder.Default
    private String importFalseString = null;

    /**
     * Specifies the separator between elements when importing as a collection.
     */
    @Builder.Default
    private String importCollectionSeparator = null;

    /**
     * Specifies whether to override a superclass field that uses the same column name on import, if one exists.
     */
    @Builder.Default
    private Boolean importOverrideSuperClassColumn = null;

    /**
     * Null-safe accessor that returns the import column names of the supplied column option, with all whitespace removed and blank entries dropped.
     *
     * @param columnOption the column option to read from; may be {@code null}
     * @return the whitespace-stripped, non-blank column names, or {@code null} if the column option or its names are {@code null}
     */
    public static List<String> getImportColumnNames(@Nullable final PxlImportColumnOption columnOption) {

        return Optional.ofNullable(columnOption)
                .flatMap(option -> Optional.ofNullable(option.getImportColumnNames())
                        .map(names -> names.stream()
                                .map(StringUtils::deleteWhitespace)
                                .filter(StringUtils::isNotBlank)
                                .collect(Collectors.toList())))
                .orElse(null);
    }

}
