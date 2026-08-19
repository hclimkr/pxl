package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import io.github.hclimkr.pxl.type.PxlExcelEngine;
import lombok.*;

import javax.validation.Valid;
import java.util.List;

/**
 * {@link CascadeWorkbook} written through the HSSF (XLS) engine.
 * <p>
 * The engine is picked after validation has already run, so the cascade decision must not depend on it. This and
 * {@link SxssfCascadeWorkbook} are the guards for that.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(exportExcelEngine = PxlExcelEngine.HSSF)
public class XlsCascadeWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @Valid
    @PxlSheet(name = "Rows")
    private List<CountingRow> rows;

}
