package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * (Guard regression) Verifies that exportOptionItems containing commas builds a dropdown (via the hidden-sheet approach) without crashing.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class OptionItemsCommaRow {

    @PxlColumn(name = "Choice", exportOptionItems = {"Apple, Inc.", "Ben & Jerry's", "AT&T"})
    private String choice;

}
