package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * An explicit inversion where importLastDataRowIndex < importFirstDataRowIndex must throw an exception at build time.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class InvertedRowWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "S", importFirstDataRowIndex = 5, importLastDataRowIndex = 3)
    private List<Employee> rows;

}
