package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import javax.validation.Valid;
import java.util.List;

/**
 * Cascading sheet with bean validation turned off on the annotation rather than through a runtime option.
 * <p>
 * Both switches end up in the same resolved workbook metadata, so this is the annotation-side twin of the
 * option-side test: {@code @Valid} on its own must trigger nothing once validation is off.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(exportDataValidation = false, importDataValidation = false)
public class NoValidationCascadeWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @Valid
    @PxlSheet(name = "Rows")
    private List<CountingRow> rows;

}
