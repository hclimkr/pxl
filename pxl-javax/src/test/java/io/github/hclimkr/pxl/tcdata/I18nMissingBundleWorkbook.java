package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook naming an i18n base name no resource answers to, for verifying that an injected bundle replaces it.
 * <p>
 * Loading the annotated base name is a {@code PxlI18nException} on its own. A bundle supplied through the option's
 * {@code import/exportResourceBundle} is used in its place before the annotated one is ever loaded, so the same
 * workbook then binds without error - which is what makes the injected bundle's precedence observable.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(
        exportI18nBaseName = "no-such-bundle",
        importI18nBaseName = "no-such-bundle")
public class I18nMissingBundleWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "staff.sheet")
    private List<I18nRow> people;

}
