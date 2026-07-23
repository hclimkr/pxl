package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlExportConverter;

/**
 * Custom object whose @PxlExportConverter returns a non-String type, used to verify
 * PxlExportConverterMeta.of rejects it (an export converter must return String).
 */
public class BadExportConverterReturnObject {

    // Invalid signature: an export converter must return String, but this returns int.
    @PxlExportConverter
    public int toExportValue() {
        return 0;
    }

}
