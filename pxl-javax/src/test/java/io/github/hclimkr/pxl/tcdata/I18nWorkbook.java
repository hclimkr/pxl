package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook for verifying i18n (import/exportI18nBaseName / Language).
 * Sheet name "people" -> "Staff", column names role/fullname -> translated to Role/Full Name.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(
        exportI18nBaseName = "messages", exportI18nLanguage = "en",
        importI18nBaseName = "messages", importI18nLanguage = "en")
public class I18nWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "people")
    private List<I18nRow> people;

}
