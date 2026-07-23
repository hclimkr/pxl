package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlImportConverter;

/**
 * Custom object whose static @PxlImportConverter (correctly returning the object type) does not take a single String,
 * used to verify PxlImportConverterMeta.of rejects it (an import converter must take exactly one String argument).
 */
public class BadImportConverterParamObject {

    // Invalid signature: an import converter must take exactly one String argument, but here it takes an Integer.
    @PxlImportConverter
    public static BadImportConverterParamObject fromString(final Integer value) {
        return new BadImportConverterParamObject();
    }

}
