package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlExportConverter;
import io.github.hclimkr.pxl.annotation.PxlImportConverter;

/**
 * Custom object for verifying that a wrongly declared @PxlExportConverter
 * (static but with 0 arguments) is caught fail-fast during column meta build.
 */
public class BadExportConverterObject {

    @PxlImportConverter
    public static BadExportConverterObject fromString(final String value) {
        return new BadExportConverterObject();
    }

    // Invalid signature: a static export converter must take 1 value argument, but here it takes 0
    @PxlExportConverter
    public static String badConverter() {
        return "x";
    }

}
