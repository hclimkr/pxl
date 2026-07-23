package io.github.hclimkr.pxl.annotation;

import java.lang.annotation.*;

/**
 * Marks a method that converts a cell string into a custom field value on import.
 * <p>
 * The annotated method is invoked by the Enum/Object codecs to turn the raw cell string into the target field
 * value, providing user-defined string-to-value mapping.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.METHOD)
@Documented
public @interface PxlImportConverter {
}
