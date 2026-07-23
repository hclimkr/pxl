package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.util.List;

/**
 * DTO for verifying custom Collection separator (collectionSeparator / import·exportCollectionSeparator) combinations.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class SeparatorRow {

    // collectionSeparator applies to both import and export.
    @PxlColumn(name = "Tags", collectionSeparator = "|")
    private List<String> tags;

    // Specify import/exportCollectionSeparator separately (same value)
    @PxlColumn(name = "Nums", importCollectionSeparator = "/", exportCollectionSeparator = "/")
    private List<Integer> nums;

}
