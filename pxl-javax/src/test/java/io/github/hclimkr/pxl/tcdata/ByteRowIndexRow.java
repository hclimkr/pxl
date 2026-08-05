package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.annotation.PxlRowIndex;
import lombok.*;

/**
 * DTO with a {@link Byte} @PxlRowIndex field, exercising the importer's Byte branch of row-index injection.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ByteRowIndexRow {

    @PxlRowIndex
    private Byte rowIndex;

    @PxlColumn(name = "Name")
    private String name;

}
