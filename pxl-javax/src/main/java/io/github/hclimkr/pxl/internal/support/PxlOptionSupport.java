package io.github.hclimkr.pxl.internal.support;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.option.PxlExportSheetOption;
import io.github.hclimkr.pxl.option.PxlImportSheetOption;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import org.apache.commons.lang3.StringUtils;

import java.util.List;
import java.util.Objects;
import java.util.function.Function;

/**
 * Lookups over the runtime option objects a caller hands to a builder.
 * <p>
 * A sheet option carries the field name it overrides, and the wildcard {@code "*"} stands for "every sheet". The
 * workbook form knows the field name it is resolving and asks for that name first, falling back to the wildcard; the
 * sheet form has no field to name and asks for the wildcard alone. Both forms are served here so the rule lives in
 * one place rather than being repeated at every site that resolves a sheet option.
 */
public final class PxlOptionSupport {

    /**
     * Prevents instantiation.
     */
    private PxlOptionSupport() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Returns the wildcard export sheet option, the only sheet-level override reachable from the sheet form.
     * <p>
     * Equivalent to {@link #findExportSheetOption(List, String)} with no field name to resolve.
     *
     * @param sheetOptions the per-field runtime sheet overrides; may be {@code null} or empty
     * @return the wildcard sheet option, or {@code null} if none is registered
     */
    public static PxlExportSheetOption findExportWildcardSheetOption(@Nullable final List<PxlExportSheetOption> sheetOptions) {

        return findExportSheetOption(sheetOptions, null);
    }

    /**
     * Returns the export sheet option that applies to the named sheet field: the option registered for that field
     * name if there is one, and the wildcard option otherwise.
     *
     * @param sheetOptions the per-field runtime sheet overrides; may be {@code null} or empty
     * @param fieldName    the sheet field name being resolved; blank or {@code null} looks up the wildcard alone
     * @return the option in effect, or {@code null} if neither is registered
     */
    public static PxlExportSheetOption findExportSheetOption(@Nullable final List<PxlExportSheetOption> sheetOptions,
                                                             @Nullable final String fieldName) {

        return findSheetOptionByFieldName(sheetOptions, fieldName, PxlExportSheetOption::getFieldName);
    }

    /**
     * Returns the wildcard import sheet option, the only sheet-level override reachable from the sheet form.
     * <p>
     * Equivalent to {@link #findImportSheetOption(List, String)} with no field name to resolve.
     *
     * @param sheetOptions the per-field runtime sheet overrides; may be {@code null} or empty
     * @return the wildcard sheet option, or {@code null} if none is registered
     */
    public static PxlImportSheetOption findImportWildcardSheetOption(@Nullable final List<PxlImportSheetOption> sheetOptions) {

        return findImportSheetOption(sheetOptions, null);
    }

    /**
     * Returns the import sheet option that applies to the named sheet field: the option registered for that field
     * name if there is one, and the wildcard option otherwise.
     *
     * @param sheetOptions the per-field runtime sheet overrides; may be {@code null} or empty
     * @param fieldName    the sheet field name being resolved; blank or {@code null} looks up the wildcard alone
     * @return the option in effect, or {@code null} if neither is registered
     */
    public static PxlImportSheetOption findImportSheetOption(@Nullable final List<PxlImportSheetOption> sheetOptions,
                                                             @Nullable final String fieldName) {

        return findSheetOptionByFieldName(sheetOptions, fieldName, PxlImportSheetOption::getFieldName);
    }

    /**
     * Resolves the option for a field name, falling back to the wildcard.
     * <p>
     * The export and import option types share no supertype, so the field name is read through {@code fieldNameGetter}
     * rather than through a common interface. The two erase to the same signature and cannot be overloads, which is
     * why the public entry points are named per direction.
     *
     * @param <T>             the option type
     * @param sheetOptions    the per-field runtime sheet overrides; may be {@code null} or empty
     * @param fieldName       the sheet field name being resolved; blank or {@code null} looks up the wildcard alone
     * @param fieldNameGetter reads the field name an option overrides
     * @return the option in effect, or {@code null} if neither the named nor the wildcard option is registered
     */
    private static <T> T findSheetOptionByFieldName(@Nullable final List<T> sheetOptions,
                                                    @Nullable final String fieldName,
                                                    final Function<T, String> fieldNameGetter) {

        if (PxlCollectionUtils.isEmpty(sheetOptions)) {
            return null;
        }

        final T namedOption = StringUtils.isBlank(fieldName)
                ? null
                : firstMatching(sheetOptions, fieldName, fieldNameGetter);

        return Objects.nonNull(namedOption)
                ? namedOption
                : firstMatching(sheetOptions, PxlConstants.SHEET_FIELD_NAME_WILD_CARD, fieldNameGetter);
    }

    /**
     * Returns the first option whose overridden field name equals {@code fieldName}.
     *
     * @param <T>             the option type
     * @param sheetOptions    the per-field runtime sheet overrides; neither {@code null} nor empty
     * @param fieldName       the field name to match
     * @param fieldNameGetter reads the field name an option overrides
     * @return the first matching option, or {@code null} if there is none
     */
    private static <T> T firstMatching(final List<T> sheetOptions,
                                       final String fieldName,
                                       final Function<T, String> fieldNameGetter) {

        return sheetOptions.stream()
                .filter(o -> StringUtils.equals(fieldNameGetter.apply(o), fieldName))
                .findFirst()
                .orElse(null);
    }

}
