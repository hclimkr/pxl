package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Control for {@link CascadeWorkbook}: the same sheet declared without {@code @Valid}, so validating the workbook
 * object stops at the workbook's own constraints and never descends into the rows. Each row is then validated
 * exactly once, by the per-row/per-collection pass.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class NoCascadeWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Rows")
    private List<CountingRow> rows;

}
