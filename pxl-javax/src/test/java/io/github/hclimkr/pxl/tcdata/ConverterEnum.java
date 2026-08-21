package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlExportConverter;
import io.github.hclimkr.pxl.annotation.PxlImportConverter;

/**
 * Enum that converts via custom @PxlExportConverter/@PxlImportConverter methods (a numeric code) rather than
 * toString/name matching, exercising the enum codec's converter-method export and import branches.
 */
public enum ConverterEnum {

    ONE,
    TWO,
    THREE,
    ;

    // String -> ConverterEnum (import path): the code is a 1-based ordinal
    @PxlImportConverter
    public static ConverterEnum fromCode(final String code) {
        return values()[Integer.parseInt(code.trim()) - 1];
    }

    // ConverterEnum -> String (export path)
    @PxlExportConverter
    public String toCode() {
        return String.valueOf(ordinal() + 1);
    }

}
