package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;

/**
 * A subclass row for verifying column override on import. (public for field access, no Lombok)
 * Because importOverrideSuperClassColumn=true, the superclass's same-named "Val" column is dropped and only the
 * subclass field is bound.
 */
public class ImportDerivedColRow extends BaseColRow {

    @PxlColumn(name = "Val", importOverrideSuperClassColumn = true)
    public String val;

}
