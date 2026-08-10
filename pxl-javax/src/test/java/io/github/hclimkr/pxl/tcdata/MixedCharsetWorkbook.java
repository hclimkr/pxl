package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import lombok.*;

import java.util.List;

/**
 * A CSV workbook whose two sheets are decoded with different charsets.
 * <p>
 * The workbook names MS949, so "Legacy" inherits it. "Modern" names UTF-8, which happens to be the built-in default -
 * re-asserting the default value against a differing workbook value is precisely what the sheet level must be able
 * to express, and it is why the annotation default is a sentinel rather than the effective default.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
@PxlWorkbook(importCsvCharset = "MS949")
public class MixedCharsetWorkbook {

    @PxlSheet(name = "Legacy")
    private List<CharsetRow> legacy;

    @PxlSheet(name = "Modern", importCsvCharset = "UTF-8")
    private List<CharsetRow> modern;

}
