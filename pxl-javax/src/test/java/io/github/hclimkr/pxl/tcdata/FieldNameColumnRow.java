package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * Row whose columns declare no name at all, so the field names stand in for them on export and on import.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class FieldNameColumnRow {

    // name is left unset ({}), so the columns are written and matched as "code" and "amount".
    @PxlColumn
    private String code;

    @PxlColumn
    private int amount;

}
