package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlExportConverter;
import io.github.hclimkr.pxl.annotation.PxlImportConverter;

/**
 * (M9-A regression) Custom object for verifying that the cause of an exception thrown by an import converter is preserved.
 */
public class ThrowingImportObject {

    @PxlImportConverter
    public static ThrowingImportObject fromString(final String value) {
        throw new IllegalStateException("converter-boom");
    }

    @PxlExportConverter
    public String toExportString() {
        return "keep";
    }

}
