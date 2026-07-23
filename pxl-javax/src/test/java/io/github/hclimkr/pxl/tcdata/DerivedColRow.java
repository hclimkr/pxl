package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;

/**
 * A subclass row for verifying column override on export.
 * Because exportOverrideSuperClassColumn=true, it overrides the superclass's same-named "Val" column so only the subclass field value is exported.
 */
public class DerivedColRow extends BaseColRow {

    @PxlColumn(name = "Val", exportOverrideSuperClassColumn = true)
    public String val;

}
