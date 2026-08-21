package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook for verifying i18n Country (export/importI18nCountry).
 * language=ko + country=KR -> Locale(ko,KR) -> resolved from messages_ko.properties (via the ko_KR -> ko candidate chain),
 * so the Korean translations (역할, 성명, 직원) are used. (differs from the English base messages.properties)
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(
        exportI18nBaseName = "messages", exportI18nLanguage = "ko", exportI18nCountry = "KR",
        importI18nBaseName = "messages", importI18nLanguage = "ko", importI18nCountry = "KR")
public class I18nCountryWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "staff.sheet")
    private List<I18nRow> people;

}
