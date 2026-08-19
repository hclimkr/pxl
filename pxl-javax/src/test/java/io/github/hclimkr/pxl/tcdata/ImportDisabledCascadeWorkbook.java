package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import javax.validation.Valid;
import java.util.List;

/**
 * Mirror of {@link DisabledCascadeWorkbook}: the {@code @Valid} sheet here is disabled for <em>import</em> and left
 * enabled for export.
 * <p>
 * The pair is what proves the sheet-cascade resolver reads the flag matching the direction rather than either flag:
 * this sheet's cascade is skipped on export (where the binder validates its rows) and kept on import (where it does
 * not), which is the exact opposite of {@link DisabledCascadeWorkbook}.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ImportDisabledCascadeWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Rows")
    private List<CountingRow> rows;

    @Valid
    @PxlSheet(name = "Skipped", importEnabled = false)
    private List<CountingRow> skippedRows;

}
