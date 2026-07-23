package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO for verifying that a header matches any of the multiple column names (name={...}) specified.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class AliasRow {

    @PxlColumn(name = {"FullName", "Name", "성명"})
    private String name;

    @PxlColumn(name = "Age")
    private int age;

}
