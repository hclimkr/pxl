package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import javax.validation.constraints.NotNull;

/**
 * Row for verifying workbook/sheet-level header stylers (required/optional).
 * req is @NotNull so the Required header styler cascades down; opt gets the Optional one.
 * (No column-level header styler is specified, so the higher-level (sheet/workbook) settings apply.)
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class HeaderStyleRow {

    @NotNull
    @PxlColumn(name = "Req")
    private String req;

    @PxlColumn(name = "Opt")
    private String opt;

}
