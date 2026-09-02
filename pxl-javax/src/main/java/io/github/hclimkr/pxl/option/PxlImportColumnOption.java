package io.github.hclimkr.pxl.option;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.support.PxlNameMatchSupport;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NonNull;

import java.util.List;
import java.util.Optional;

/**
 * Column import option, for an Excel or a CSV source alike.
 */
@Getter
@AllArgsConstructor
@Builder
public final class PxlImportColumnOption {

    /**
     * The field name for the column object
     */
    @NonNull    // required
    private final String fieldName;

    /**
     * Specifies the name of the column on import.
     * <p>
     * Doubles as a content-i18n key, just like the {@link PxlColumn#name()} it overrides:
     * when the workbook sets {@code importI18nBaseName}, the name is resolved through that bundle first and it is the
     * translation that the header is matched against. A name the bundle does not carry is used as it stands.
     */
    @Builder.Default
    private final List<String> importColumnNames = null;

    /**
     * Specifies whether to import.
     */
    @Builder.Default
    private final Boolean importEnabled = null;

    /**
     * Specifies whether to trim strings on import.
     */
    @Builder.Default
    private final Boolean importTrim = null;

    /**
     * Specifies whether to check the uniqueness of the column's values on import.
     */
    @Builder.Default
    private final Boolean importUnique = null;

    /**
     * Specifies the cell formatting string on import.
     * Valid only for fields of type Numeric, Date, LocalTime, LocalDate, LocalDateTime, ZonedDateTime, OffsetTime, OffsetDateTime, Duration, and Period.
     * A Duration/Period pattern is the {@code DurationFormatUtils} style used on export, and a value that does not match it falls back to ISO-8601.
     * The pattern has to consume the cell value in full; a value it reads only the front of is rejected.
     */
    @Builder.Default
    private final String importPattern = null;

    /**
     * Specifies the string that represents the boolean value true on import.
     * <p>
     * A String column renders a BOOLEAN cell as this string, and a Boolean column interprets this string
     * (case-insensitive) as true.
     */
    @Builder.Default
    private final String importTrueString = null;

    /**
     * Specifies the string that represents the boolean value false on import.
     * <p>
     * A String column renders a BOOLEAN cell as this string, and a Boolean column interprets this string
     * (case-insensitive) as false.
     */
    @Builder.Default
    private final String importFalseString = null;

    /**
     * Specifies the separator between elements when importing as a collection.
     */
    @Builder.Default
    private final String importCollectionSeparator = null;

    /**
     * Specifies whether to override a superclass field that uses the same column name on import, if one exists.
     */
    @Builder.Default
    private final Boolean importOverrideSuperClassColumn = null;

    /**
     * Null-safe accessor that returns the import column names of the supplied column option, with all whitespace removed and blank entries dropped.
     *
     * @param columnOption the column option to read from; may be {@code null}
     * @return the whitespace-stripped, non-blank column names, or {@code null} if the column option or its names are {@code null}
     */
    public static List<String> getImportColumnNames(@Nullable final PxlImportColumnOption columnOption) {

        return Optional.ofNullable(columnOption)
                .flatMap(option -> Optional.ofNullable(option.getImportColumnNames())
                        .map(PxlNameMatchSupport::normalizeNames))
                .orElse(null);
    }

}
