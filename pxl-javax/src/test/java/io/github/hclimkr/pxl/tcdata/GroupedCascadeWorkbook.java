package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import javax.validation.Valid;
import java.util.List;

/**
 * Grouping export ({@code exportGroupingFieldName}) of a {@code @Valid} sheet.
 * <p>
 * Grouping redistributes the rows into one sheet per key before writing, so the export walks a different code path
 * than a plain sheet does. Validation still happens once over the original collection, and the cascade still has to
 * be skipped - each row counted exactly once no matter how many sheets it ends up spread across.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class GroupedCascadeWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @Valid
    @PxlSheet(name = "Rows", exportGroupingFieldName = "age")
    private List<CountingRow> rows;

}
