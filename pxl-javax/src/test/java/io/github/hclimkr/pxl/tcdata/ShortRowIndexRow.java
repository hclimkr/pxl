package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.annotation.PxlRowIndex;
import lombok.*;

/**
 * DTO with a {@code Short} @PxlRowIndex field, exercising the importer's Short branch of row-index injection.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ShortRowIndexRow {

    @PxlRowIndex
    private Short rowIndex;

    @PxlColumn(name = "Name")
    private String name;

}
