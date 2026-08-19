package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import io.github.hclimkr.pxl.type.PxlExcelEngine;
import lombok.*;

import javax.validation.Valid;
import java.util.List;

/**
 * {@link CascadeWorkbook} written through the SXSSF (streaming XLSX) engine.
 * <p>
 * Streaming flushes rows to disk as it goes, but validation has already finished by then, so the cascade decision
 * must come out the same as with {@link XlsCascadeWorkbook} and the in-memory default.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(exportExcelEngine = PxlExcelEngine.SXSSF)
public class SxssfCascadeWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @Valid
    @PxlSheet(name = "Rows")
    private List<CountingRow> rows;

}
