package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * (M9-A regression) A custom-object column whose import converter throws an exception.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ThrowingConverterRow {

    @PxlColumn(name = "Value")
    private ThrowingImportObject value;

}
