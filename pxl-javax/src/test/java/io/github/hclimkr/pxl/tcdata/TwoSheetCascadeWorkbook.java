package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import javax.validation.Valid;
import java.util.List;

/**
 * Two sheets that are both enabled, only one of them cascading.
 * <p>
 * Every other multi-sheet fixture here has one sheet disabled, which makes the two halves behave differently for
 * that reason. This one isolates the remaining question: with both sheets processed normally, each sheet's rows are
 * validated exactly once whether or not that sheet carries {@code @Valid}.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class TwoSheetCascadeWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @Valid
    @PxlSheet(name = "First")
    private List<CountingRow> firstRows;

    @PxlSheet(name = "Second")
    private List<CountingRow> secondRows;

}
