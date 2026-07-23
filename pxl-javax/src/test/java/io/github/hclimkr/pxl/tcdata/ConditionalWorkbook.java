package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import lombok.*;

import java.util.List;

/**
 * Workbook for verifying exportIfNull / exportIfEmpty (whether a sheet is created when the data is null/empty list).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ConditionalWorkbook {

    // create the sheet even when null
    @PxlSheet(name = "KeepWhenNull", exportIfNull = true)
    private List<Employee> keepWhenNull;

    // no sheet created when null (exportIfNull defaults to false)
    @PxlSheet(name = "DropWhenNull")
    private List<Employee> dropWhenNull;

    // no sheet created when empty
    @PxlSheet(name = "DropWhenEmpty", exportIfEmpty = false)
    private List<Employee> dropWhenEmpty;

    // create the sheet even when empty (exportIfEmpty defaults to true)
    @PxlSheet(name = "KeepWhenEmpty")
    private List<Employee> keepWhenEmpty;

}
