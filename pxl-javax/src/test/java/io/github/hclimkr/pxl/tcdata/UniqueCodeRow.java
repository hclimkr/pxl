package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO for verifying importUnique (uniqueness check of column values). If code is duplicated, import throws an exception.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class UniqueCodeRow {

    @PxlColumn(name = "Code", importUnique = true)
    private String code;

    @PxlColumn(name = "Name")
    private String name;

}
