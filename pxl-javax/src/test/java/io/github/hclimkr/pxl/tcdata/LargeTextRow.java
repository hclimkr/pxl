package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO for driving a CSV export past the size at which it stops fitting in memory. One wide text column makes the
 * output large from few rows, and a Double column gives the export a value the codec can reject, so a failure can
 * be placed after the render has already spilled to a temporary file.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class LargeTextRow {

    @PxlColumn(name = "Text")
    private String text;

    @PxlColumn(name = "Value")
    private Double value;

}
