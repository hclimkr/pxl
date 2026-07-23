package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * (M8 regression) Verifies that leading/trailing whitespace in a column name derived from the annotation is trimmed in the export header.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class SpacedNameRow {

    @PxlColumn(name = "  Padded Name  ")
    private String value;

}
