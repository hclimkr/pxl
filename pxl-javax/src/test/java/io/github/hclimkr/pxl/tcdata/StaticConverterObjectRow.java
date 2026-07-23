package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO with a {@link StaticConverterObject} column, for round-tripping a custom object whose export converter is static.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class StaticConverterObjectRow {

    @PxlColumn(name = "Value")
    private StaticConverterObject value;

}
