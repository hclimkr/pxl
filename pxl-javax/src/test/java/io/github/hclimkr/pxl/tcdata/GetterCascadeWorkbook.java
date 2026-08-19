package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;

import javax.validation.Valid;
import java.util.List;

/**
 * Declares the cascade on the getter rather than on the field, which switches Bean Validation to property access
 * and makes the provider report the traversal with {@code ElementType.METHOD}.
 * <p>
 * {@code @PxlSheet} itself is field-only, so the two annotations sit on different members here. The resolver
 * matches on the property name, which is the same either way - this fixture is what proves that holds rather than
 * being an accident of field access.
 * <p>
 * Written without Lombok because {@code @Getter} would generate the very accessor that has to carry
 * {@code @Valid}.
 */
public class GetterCascadeWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Rows")
    private List<CountingRow> rows;

    public String getWorkbookName() {

        return workbookName;
    }

    public void setWorkbookName(final String workbookName) {

        this.workbookName = workbookName;
    }

    @Valid
    public List<CountingRow> getRows() {

        return rows;
    }

    public void setRows(final List<CountingRow> rows) {

        this.rows = rows;
    }

}
