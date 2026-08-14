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
     * The message template reported when the byte length falls outside the range.
     *
     * @return the error message template used when the constraint is violated
     */
    String message() default "byte length is out of the allowed range ({min}~{max}).";

    /**
     * The validation groups that have to be active for this constraint to be checked.
     *
     * @return the validation groups this constraint belongs to
     */
    Class<?>[] groups() default {};

    /**
     * The metadata carried alongside this constraint, for a client of the validation API to read.
     *
     * @return the payload associated with this constraint
     */
    Class<? extends Payload>[] payload() default {};

    /**
     * The smallest byte length the element may have.
     *
     * @return byte length the element must be higher or equal to
     */
    int min() default 0;

    /**
     * The largest byte length the element may have.
     *
     * @return byte length the element must be lower or equal to
     */
    int max() default Integer.MAX_VALUE;

    /**
     * The charset the value is encoded with before its bytes are counted, which is what makes the limit depend on
     * the encoding rather than on the character count.
     *
     * @return the charset name the value is encoded with before its bytes are counted; {@code "UTF-8"} by default
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
         * The constraints that apply to the element, one per declared annotation.
         *
         * @return the {@link PxlByteSize} annotations declared on the element
         */
        PxlByteSize[] value();
    }

}
