package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import lombok.*;

import java.util.List;

/**
 * A workbook whose CSV charset/delimiter are deliberately unusable, for verifying that an Excel source ignores them.
 * <p>
 * Neither value could ever open a CSV — the charset names nothing this JVM carries, and the delimiter is the quote
 * character CSVFormat rejects. An Excel import must nevertheless succeed, since it reads one file whose encoding the
 * format itself carries and which has no delimiter at all.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(importCsvCharset = "NoSuchCharset-1")
public class IgnoredCsvAttrsWorkbook {

    @PxlSheet(name = "Employees", importCsvDelimiter = '"')
    private List<Employee> employees;

}
