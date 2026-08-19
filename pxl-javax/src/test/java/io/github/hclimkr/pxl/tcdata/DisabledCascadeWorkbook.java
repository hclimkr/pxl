package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import javax.validation.Valid;
import java.util.List;

/**
 * Workbook with one exported sheet and one sheet turned off for export, the disabled one carrying {@code @Valid}.
 * <p>
 * A disabled sheet is skipped by the binder, so its rows never reach the per-row validation pass. Skipping its
 * cascade as well would drop those rows from validation entirely, so the sheet-cascade resolver has to leave a
 * disabled sheet alone - that is what this fixture pins down.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class DisabledCascadeWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Rows")
    private List<CountingRow> rows;

    @Valid
    @PxlSheet(name = "Skipped", exportEnabled = false)
    private List<CountingRow> skippedRows;

}
