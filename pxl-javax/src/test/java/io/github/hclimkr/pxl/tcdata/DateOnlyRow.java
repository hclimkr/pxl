package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.util.Date;

/**
 * (Regression for the Date seconds-loss fix) Verifies that a {@link Date} with no pattern specified round-trips down to the second under default settings.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class DateOnlyRow {

    @PxlColumn(name = "When")
    private Date when;

}
