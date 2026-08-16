package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO for verifying which form a String column is written in when its value starts with {@code '='}.
 * Both columns are declared in a fixed order so a test can address them by index (issue L7).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class PicturePrecedenceRow {

    // Picture only: a leading '=' must not divert the value into the text path
    @PxlColumn(name = "PictureOnly", exportOrder = "1", exportStringAsPicture = true)
    private String pictureOnly;

    // Both options: the formula wins
    @PxlColumn(name = "PictureAndFormula", exportOrder = "2", exportStringAsPicture = true, exportStringAsFormula = true)
    private String pictureAndFormula;

}
