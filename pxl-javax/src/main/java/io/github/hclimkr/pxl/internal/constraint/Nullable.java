package io.github.hclimkr.pxl.internal.constraint;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.*;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

/**
 * Marks the annotated element (method return value, field, parameter, or local variable) as possibly {@code null}.
 * <p>
 * This is a documentation and static-analysis marker only; it is not enforced at runtime. It exists so PXL can
 * express nullability without depending on a platform-specific package (the {@code javax.annotation} /
 * {@code jakarta.annotation} split), keeping a single source tree portable across both platforms.
 */
@Documented
@Retention(RUNTIME)
@Target({METHOD, FIELD, PARAMETER, LOCAL_VARIABLE, TYPE_USE})
public @interface Nullable {
}
