package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO whose exportSample values exercise the formula and picture export forms.
 * A sample value is a raw string handed to the same codec a data value goes through, so both options apply to the
 * sample row as well - the columns are ordered so a test can address them by index.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class FormulaPictureSampleRow {

    @PxlColumn(name = "Formula", exportOrder = "1", exportStringAsFormula = true, exportSample = "=1+2")
    private String formula;

    @PxlColumn(name = "Photo", exportOrder = "2", exportStringAsPicture = true, exportSample = "photo.png")
    private String photo;

}
