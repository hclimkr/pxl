package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;

import javax.validation.Valid;
import java.util.List;

/**
 * Subclass shadowing {@link SuperShadowCascadeWorkbook}'s sheet field with an enabled declaration of the same name.
 * <p>
 * See the parent for what this pair pins down: a field name carrying both an enabled and a disabled declaration
 * must keep its cascade.
 */
public class SubShadowCascadeWorkbook extends SuperShadowCascadeWorkbook {

    @Valid
    @PxlSheet(name = "Rows")
    public List<CountingRow> rows;

}
