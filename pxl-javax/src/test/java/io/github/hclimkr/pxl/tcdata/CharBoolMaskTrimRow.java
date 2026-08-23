package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO with {@code char}, {@link Character} and {@link Boolean} columns carrying {@code exportMasking} or
 * {@code exportTrim}, so each rendered value goes through the shared string-level export processing.
 * The boolean trim column pads its true/false strings, because "true"/"false" have no whitespace to trim.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class CharBoolMaskTrimRow {

    @PxlColumn(name = "MaskWrapChar", exportMasking = "[a-z]")
    private Character maskWrapChar;

    @PxlColumn(name = "MaskPrimChar", exportMasking = "[a-z]")
    private char maskPrimChar;

    @PxlColumn(name = "MaskBool", exportMasking = "[a-z]")
    private Boolean maskBool;

    @PxlColumn(name = "TrimWrapChar", exportTrim = true)
    private Character trimWrapChar;

    @PxlColumn(name = "TrimPrimChar", exportTrim = true)
    private char trimPrimChar;

    @PxlColumn(name = "TrimBool", exportTrim = true, exportTrueString = " Y ", exportFalseString = " N ")
    private Boolean trimBool;

}
