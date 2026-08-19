package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import javax.validation.Valid;
import java.util.List;

/**
 * {@link CascadeWorkbook} read through the streaming reader instead of the in-memory one.
 * <p>
 * The streaming path is a different reader (StreamingWorkbook) but the same binder, so the sheet cascade has to be
 * skipped there too - this fixture is what keeps the two import paths from drifting apart.
 * <p>
 * The streaming reader does not support getFirstRowNum(), so the sheet states its header and first data row
 * explicitly (1-based).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(importUsingStreamReader = true)
public class StreamCascadeWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @Valid
    @PxlSheet(name = "Rows", importHeaderRowIndex = 1, importFirstDataRowIndex = 2)
    private List<CountingRow> rows;

}
