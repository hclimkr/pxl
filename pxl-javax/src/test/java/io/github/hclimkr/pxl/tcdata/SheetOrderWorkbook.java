package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Verifies sheet exportOrder. The names are Zebra/Apple, but according to exportOrder (A,B) Zebra must come first.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class SheetOrderWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Zebra", exportOrder = "A")
    private List<Employee> zebra;

    @PxlSheet(name = "Apple", exportOrder = "B")
    private List<Employee> apple;

}
