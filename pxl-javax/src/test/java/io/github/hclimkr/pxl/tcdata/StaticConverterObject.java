package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlExportConverter;
import io.github.hclimkr.pxl.annotation.PxlImportConverter;

/**
 * Custom object whose export converter is a STATIC @PxlExportConverter method (unlike {@link Money}, whose
 * export converter is an instance method), exercising the object codec's static-export-converter branch.
 * The private constructor keeps the String-constructor import path out, so import goes through the static
 *
 * @PxlImportConverter.
 */
public class StaticConverterObject {

    private final String value;

    private StaticConverterObject(final String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }

    // String -> StaticConverterObject (import path): static converter method
    @PxlImportConverter
    public static StaticConverterObject fromString(final String value) {
        return new StaticConverterObject(value);
    }

    // StaticConverterObject -> String (export path): STATIC converter method (takes the object)
    @PxlExportConverter
    public static String toStaticString(final StaticConverterObject object) {
        return object.value;
    }

}
