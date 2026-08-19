package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import javax.validation.Valid;
import javax.validation.constraints.NotBlank;
import javax.validation.constraints.NotEmpty;
import java.util.List;

/**
 * Workbook form that pins down what skipping the sheet cascade must NOT take away: a plain constraint on the
 * workbook's own field ({@code @NotBlank}), a constraint on the sheet collection itself ({@code @NotEmpty}) and a
 * {@code @Valid} on a field that is not a sheet all have to keep working, while the rows behind {@code @Valid} are
 * left to the per-row validation pass.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ConstrainedCascadeWorkbook {

    @NotBlank(message = "'WorkbookName' must not be blank.")
    @PxlWorkbookName
    private String workbookName;

    @Valid
    @NotEmpty(message = "'Rows' must not be empty.")
    @PxlSheet(name = "Rows")
    private List<CountingRow> rows;

    @Valid
    private CascadeMeta meta;

}
