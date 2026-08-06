package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import lombok.*;

import java.util.List;

/**
 * A CSV workbook carrying @PxlWorkbook but naming neither a charset nor a delimiter.
 * <p>
 * Both annotation levels then hold the "not specified" sentinel, so the resolved values must come from the built-in
 * defaults (UTF-8 and the comma). Reading the sentinel as a usable value instead would hand {@code ""} to
 * {@code Charset.forName}, which is why this is the workbook the cascade is at most risk of breaking.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook
public class DefaultCsvWorkbook {

    @PxlSheet(name = "Cities")
    private List<CharsetRow> cities;

    @PxlSheet(name = "Departments")
    private List<Department> departments;

}
