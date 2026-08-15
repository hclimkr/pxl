package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook that asks for the streaming reader, with its row cache and buffer sized, on the annotation rather than in
 * a runtime option.
 * <p>
 * The streaming reader does not support getFirstRowNum(), so the sheet states its header and first data row
 * explicitly (1-based).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(importUsingStreamReader = true, importStreamReaderRowCacheSize = 50, importStreamReaderBufferSize = 8192)
public class StreamReaderWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "People", importHeaderRowIndex = 1, importFirstDataRowIndex = 2)
    private List<Employee> people;

}
