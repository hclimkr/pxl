package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.util.List;

/**
 * DTO whose Collection column carries a DecimalFormat pattern, so the pattern is applied to each element rather than
 * to the joined string. For verifying that an element the pattern cannot consume in full is rejected (issue M3).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class CollectionPatternRow {

    @PxlColumn(name = "Nums", pattern = "#,##0")
    private List<Integer> nums;

}
