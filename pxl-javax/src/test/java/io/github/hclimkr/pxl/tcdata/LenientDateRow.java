package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.time.LocalDate;

/**
 * For verifying that importPattern is not strictly enforced and falls back to the default pattern (import asymmetry).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class LenientDateRow {

    @PxlColumn(name = "D", importPattern = "dd/MM/yyyy")
    private LocalDate d;

}
