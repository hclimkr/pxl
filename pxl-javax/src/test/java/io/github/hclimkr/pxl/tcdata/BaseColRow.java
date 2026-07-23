package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;

/**
 * Parent row for verifying column override on export. (public for field access, no Lombok)
 */
public class BaseColRow {

    @PxlColumn(name = "Val")
    public String val;

}
