package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO with a {@link ConverterEnum} column, for round-tripping an enum through its custom export/import converters.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ConverterEnumRow {

    @PxlColumn(name = "Code")
    private ConverterEnum code;

}
