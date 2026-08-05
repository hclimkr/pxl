package io.github.hclimkr.pxl.annotation;

import io.github.hclimkr.pxl.exception.PxlDataException;

import java.lang.annotation.*;

/**
 * Marks the field that receives the workbook name on import (Excel and CSV alike, in the workbook form).
 * <p>
 * The field is filled with the name given to the import builder's {@code workbookName(...)}. When no name is given,
 * an Excel import from a file names the workbook after the source file, with the extension removed; a stream source
 * carries no file name, so the field is left untouched there.
 * <p>
 * The annotated field must be of type {@link String}; otherwise a {@link PxlDataException} is raised — the type is
 * checked while collecting metadata, on the export path as well, even though export never reads the field. Only the
 * first such field is used, scanning the class and then its superclasses.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface PxlWorkbookName {
}
