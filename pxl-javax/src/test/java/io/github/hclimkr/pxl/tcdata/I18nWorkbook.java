package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook for verifying i18n (export/importI18nBaseName / Language).
 * Sheet key staff.sheet -> "Staff", column keys staff.column.role/fullName -> translated to Role/Full Name.
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

    @PxlSheet(name = "staff.sheet")
    private List<I18nRow> people;

}
