package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO for verifying exportNullString (the string shown when a value is null).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class NullStringRow {

    @PxlColumn(name = "Value", exportNullString = "N/A")
    private String value;

    @PxlColumn(name = "Label")
    private String label;

}
