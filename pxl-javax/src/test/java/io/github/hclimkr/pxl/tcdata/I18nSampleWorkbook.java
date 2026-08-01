package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook for verifying content i18n of the sample row (exportSample / exportOptionItems).
 * Sheet key staff.sheet -> "Staff"; the bundle is the same messages.properties the other i18n fixtures use.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(
        exportI18nBaseName = "messages", exportI18nLanguage = "en",
        importI18nBaseName = "messages", importI18nLanguage = "en")
public class I18nSampleWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "staff.sheet")
    private List<I18nSampleRow> people;

}
