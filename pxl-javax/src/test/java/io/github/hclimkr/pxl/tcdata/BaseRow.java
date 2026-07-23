package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.experimental.SuperBuilder;

/**
 * Parent row DTO for verifying column inheritance. Subclasses inherit these @PxlColumn fields.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@SuperBuilder
public class BaseRow {

    @PxlColumn(name = "Id")
    private Integer id;

    @PxlColumn(name = "BaseName")
    private String baseName;

}
