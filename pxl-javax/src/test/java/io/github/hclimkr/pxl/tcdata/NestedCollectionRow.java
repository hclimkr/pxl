package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.util.List;

/**
 * Nested generic Collections are not supported, so an exception must be thrown.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class NestedCollectionRow {

    @PxlColumn(name = "Nested")
    private List<List<String>> nested;

}
