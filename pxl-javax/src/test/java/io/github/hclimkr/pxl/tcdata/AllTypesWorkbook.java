package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * A @PxlWorkbook-style DTO holding all-types rows (AllTypesRow). For workbook-shape roundtrip verification.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class AllTypesWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "AllTypes")
    private List<AllTypesRow> rows;

}
