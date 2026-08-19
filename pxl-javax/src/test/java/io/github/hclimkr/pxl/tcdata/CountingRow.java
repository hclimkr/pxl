package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * Row DTO whose only constraint counts its own evaluations, so a test can tell how many times bean validation
 * visited the row. Paired with {@link CascadeWorkbook} / {@link NoCascadeWorkbook}, which declare the same sheet
 * with and without {@code @Valid}.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class CountingRow {

    @CountedNotBlank(message = "'Name' must not be blank.")
    @PxlColumn(name = "Name")
    private String name;

    @PxlColumn(name = "Age")
    private Integer age;

}
