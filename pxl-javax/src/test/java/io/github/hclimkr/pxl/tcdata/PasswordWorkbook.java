package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook that carries the export and import passwords on the annotation rather than in a runtime option.
 * The two match, so the same class both protects the document and opens it again.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(exportPassword = "secret", importPassword = "secret")
public class PasswordWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "People")
    private List<Employee> people;

}
