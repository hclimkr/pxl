package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO with a column of an unsupported type. Meta build must throw an exception.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class UnsupportedTypeRow {

    @PxlColumn(name = "U")
    private Unsupported u;

}
