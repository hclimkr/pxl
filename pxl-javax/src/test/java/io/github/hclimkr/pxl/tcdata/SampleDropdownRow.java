package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO for verifying that a sample export sizes its auto-filter and dropdown ranges to the one sample row it
 * writes, rather than to a declared exportLastDataRowIndex. Carries a plain sampled column and a sampled column
 * with a fixed dropdown list, so both the filter range and the data validation can be asserted.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class SampleDropdownRow {

    @PxlColumn(name = "Name", exportSample = "Alice")
    private String name;

    @PxlColumn(name = "Choice", exportSample = "Red", exportOptionItems = {"Red", "Green", "Blue"})
    private String choice;

}
