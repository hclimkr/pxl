package io.github.hclimkr.pxl.constraint;

import io.github.hclimkr.pxl.internal.constraint.PxlByteSizeValidator;

import javax.validation.Constraint;
import javax.validation.Payload;

import java.lang.annotation.Documented;
import java.lang.annotation.Repeatable;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.*;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

/**
 * The annotated element byte length must be between the specified boundaries (included).
 * <p>
 * Supported types are:
 * <ul>
 *     <li>{@link CharSequence} (byte length of character sequence is evaluated)</li>
 * </ul>
 * <p>
 * {@code null} elements are considered valid. Determine the byte length by encoding the string in the specified
 * {@link PxlByteSize#charset()}. If not specified, the string is encoded with charset {@code "UTF-8"}.
 */
@Target({METHOD, FIELD, ANNOTATION_TYPE, CONSTRUCTOR, PARAMETER, TYPE_USE})
@Retention(RUNTIME)
@Repeatable(PxlByteSize.List.class)
@Documented
@Constraint(validatedBy = {PxlByteSizeValidator.class})
public @interface PxlByteSize {

    /**
     * @return the error message template used when the constraint is violated
     */
    String message() default "byte length is out of the allowed range ({min}~{max}).";

    /**
     * @return the validation groups this constraint belongs to
     */
    Class<?>[] groups() default {};

    /**
     * @return the payload associated with this constraint
     */
    Class<? extends Payload>[] payload() default {};

    /**
     * @return byte length the element must be higher or equal to
     */
    int min() default 0;

    /**
     * @return byte length the element must be lower or equal to
     */
    int max() default Integer.MAX_VALUE;

    /**
     * @return the charset name used in parse to a string
     */
    String charset() default "UTF-8";

    /**
     * Defines several {@link PxlByteSize} annotations on the same element.
     *
     * @see PxlByteSize
     */
    @Target({METHOD, FIELD, ANNOTATION_TYPE, CONSTRUCTOR, PARAMETER, TYPE_USE})
    @Retention(RUNTIME)
    @Documented
    @interface List {

        /**
         * @return the {@link PxlByteSize} annotations declared on the element
         */
        PxlByteSize[] value();
    }

}
