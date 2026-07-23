package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.annotation.PxlRowIndex;
import lombok.*;

/**
 * DTO with a {@code Long} @PxlRowIndex field, exercising the importer's Long branch of row-index injection.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class LongRowIndexRow {

    @PxlRowIndex
    private Long rowIndex;

    @PxlColumn(name = "Name")
    private String name;

}
