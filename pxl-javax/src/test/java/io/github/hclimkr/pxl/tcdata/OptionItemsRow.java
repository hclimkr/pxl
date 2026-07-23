package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO for verifying exportOptionItems (fixed-list dropdown).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class OptionItemsRow {

    @PxlColumn(name = "Choice", exportOptionItems = {"Red", "Green", "Blue"})
    private String choice;

}
