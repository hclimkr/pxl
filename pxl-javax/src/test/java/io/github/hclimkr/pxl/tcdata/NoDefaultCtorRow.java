package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * DTO without a no-arg constructor (only an all-args constructor). Reflective instantiation on import must fail.
 */
@Getter
@AllArgsConstructor
public class NoDefaultCtorRow {

    @PxlColumn(name = "Name")
    private String name;

}
