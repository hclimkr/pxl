package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook whose sheet declares no name at all, so the field name stands in for it on export and on import.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class FieldNameSheetWorkbook {

    @PxlWorkbookName
    private String workbookName;

    // name is left unset ({}), so the sheet is written and matched as "employees".
    @PxlSheet
    private List<Employee> employees;

}
