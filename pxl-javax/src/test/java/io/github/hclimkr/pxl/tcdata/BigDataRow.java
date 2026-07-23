package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * Lightweight row DTO for large-data streaming tests. (2 columns to minimize build/parse cost)
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class BigDataRow {

    @PxlColumn(name = "Id")
    private int id;

    @PxlColumn(name = "Name")
    private String name;

}
