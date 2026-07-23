package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.experimental.SuperBuilder;

/**
 * A subclass row DTO for verifying column inheritance. Both the superclass (BaseRow) columns and its own columns must be exported/imported.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@SuperBuilder
public class DerivedRow extends BaseRow {

    @PxlColumn(name = "Extra")
    private String extra;

}
