package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import lombok.*;

import java.util.List;

/**
 * A CSV workbook whose two sheets are split on different delimiters.
 * <p>
 * The workbook names the tab, so "Tabbed" inherits it. "Comma" names the comma, which happens to be the built-in
 * default - the delimiter counterpart of what {@link MixedCharsetWorkbook} does for the charset.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(importCsvDelimiter = '\t')
public class MixedDelimiterWorkbook {

    @PxlSheet(name = "Tabbed")
    private List<Department> tabbed;

    @PxlSheet(name = "Comma", importCsvDelimiter = ',')
    private List<Department> comma;

}
