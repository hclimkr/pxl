package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.annotation.PxlRowIndex;
import lombok.*;

/**
 * DTO that applies @PxlRowIndex to an unsupported type (String). An exception must be thrown on import.
 * (Supported types: byte/short/int/long and their wrapper classes Byte/Short/Integer/Long)
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class BadRowIndexRow {

    @PxlRowIndex
    private String rowIndex;   // invalid type

    @PxlColumn(name = "Name")
    private String name;

}
