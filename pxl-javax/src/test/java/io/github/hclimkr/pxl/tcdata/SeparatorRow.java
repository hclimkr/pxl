package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.util.List;

/**
 * DTO for verifying custom Collection separator (collectionSeparator / export/importCollectionSeparator) combinations.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class SeparatorRow {

    // collectionSeparator applies to both export and import.
    @PxlColumn(name = "Tags", collectionSeparator = "|")
    private List<String> tags;

    // Specify export/importCollectionSeparator separately (same value)
    @PxlColumn(name = "Nums", exportCollectionSeparator = "/", importCollectionSeparator = "/")
    private List<Integer> nums;

}
