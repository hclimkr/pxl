package io.github.hclimkr.pxl.annotation;

import java.lang.annotation.*;

/**
 * Marks the field that receives the Excel workbook name on import.
 * <p>
 * The annotated field must be of type {@link String}; otherwise a {@code PxlDataException} is raised. Only the
 * first such field (including inherited fields) is used.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface PxlWorkbookName {
}
