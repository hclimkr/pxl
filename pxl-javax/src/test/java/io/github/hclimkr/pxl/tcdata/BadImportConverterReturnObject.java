package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlImportConverter;

/**
 * Custom object whose static @PxlImportConverter returns a type other than the object type, used to verify
 * PxlImportConverterMeta.of rejects it (an import converter must return the object type).
 */
public class BadImportConverterReturnObject {

    // Invalid signature: an import converter must return BadImportConverterReturnObject, but this returns String.
    @PxlImportConverter
    public static String fromString(final String value) {
        return value;
    }

}
