package io.github.hclimkr.pxl.annotation;

import java.lang.annotation.*;

/**
 * Marks a method that converts a custom field value into a cell string on export.
 * <p>
 * The annotated method is invoked by the Enum/Object codecs to turn the field value into the string written to
 * the cell, providing user-defined value-to-string mapping.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.METHOD)
@Documented
public @interface PxlExportConverter {
}
