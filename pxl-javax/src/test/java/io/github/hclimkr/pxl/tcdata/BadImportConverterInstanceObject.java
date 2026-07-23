package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlImportConverter;

/**
 * Custom object whose @PxlImportConverter (correctly returning the object type) is an instance method, used to verify
 * PxlImportConverterMeta.of rejects it (an import converter must be static).
 */
public class BadImportConverterInstanceObject {

    // Invalid signature: an import converter must be static, but this is an instance method.
    @PxlImportConverter
    public BadImportConverterInstanceObject fromString(final String value) {
        return this;
    }

}
