package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * For verifying that when a pattern is specified, a long is recorded as text so precision is preserved even beyond 2^53.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class LongPatternRow {

    @PxlColumn(name = "Big", pattern = "0")
    private Long big;

}
