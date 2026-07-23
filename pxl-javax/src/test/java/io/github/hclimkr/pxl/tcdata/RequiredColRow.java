package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import javax.validation.constraints.NotBlank;

/**
 * Verifies an exception when a required column (@NotBlank) header is missing, and null when an optional column is missing.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class RequiredColRow {

    @NotBlank
    @PxlColumn(name = "Req")
    private String req;

    @PxlColumn(name = "Opt")
    private String opt;

}
