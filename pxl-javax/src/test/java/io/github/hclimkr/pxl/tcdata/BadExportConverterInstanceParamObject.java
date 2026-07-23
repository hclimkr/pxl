package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlExportConverter;

/**
 * Custom object whose instance @PxlExportConverter (correctly returning String) takes an argument, used to verify
 * PxlExportConverterMeta.of rejects it (an instance export converter must take no argument).
 */
public class BadExportConverterInstanceParamObject {

    // Invalid signature: an instance export converter must take 0 arguments, but here it takes 1.
    @PxlExportConverter
    public String toExportString(final int unused) {
        return "x";
    }

}
