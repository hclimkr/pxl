package io.github.hclimkr.pxl.annotation;

import java.lang.annotation.*;

/**
 * Marks the field that receives the source row index on import.
 * <p>
 * The value written is the 1-based spreadsheet row number of the imported row (the underlying 0-based POI row
 * number plus one; e.g. with the header on the first row, the first data row receives {@code 2}). The annotated field must be
 * one of {@code byte}, {@code short}, {@code int}, {@code long} or their wrapper classes ({@code Byte},
 * {@code Short}, {@code Integer}, {@code Long}); otherwise a
 * {@code PxlArgumentException} is raised. Only the first such field (including inherited fields) is used.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface PxlRowIndex {
}   
