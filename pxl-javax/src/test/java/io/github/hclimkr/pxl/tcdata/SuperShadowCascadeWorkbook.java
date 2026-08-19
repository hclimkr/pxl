package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;

import javax.validation.Valid;
import java.util.List;

/**
 * Parent of {@link SubShadowCascadeWorkbook}, declaring the sheet field <em>disabled</em> for export.
 * <p>
 * The subclass shadows the same field name with an enabled declaration, so one field name resolves to two
 * {@code @PxlSheet} declarations that disagree. The sheet-cascade resolver reads field names, not resolved sheet
 * metadata, so it cannot tell which one wins - and it keeps the cascade when any declaration is disabled, since
 * validating rows twice is recoverable and not validating them is not.
 * <p>
 * Uses public fields and no Lombok, like the other shadowing fixtures: a setter would be overridden and could no
 * longer address the two fields separately.
 */
public class SuperShadowCascadeWorkbook {

    @PxlWorkbookName
    public String workbookName;

    @Valid
    @PxlSheet(name = "Rows", exportEnabled = false)
    public List<CountingRow> rows;

}
