package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook that turns bean validation off on both directions through the annotation, so a row violating its own
 * constraints travels out and back without an exception.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(exportDataValidation = false, importDataValidation = false)
public class NoValidationWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Rows")
    private List<ValidatedRow> rows;

}
