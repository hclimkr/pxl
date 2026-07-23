package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * For verifying CSV encoding (importCsvCharset). Reads non-ASCII characters with a specific charset.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class CharsetRow {

    @PxlColumn(name = "City")
    private String city;

}
