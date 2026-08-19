package io.github.hclimkr.pxl.tcdata;

import javax.validation.Constraint;
import javax.validation.Payload;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

/**
 * Test-only constraint that rejects a blank value just like {@code @NotBlank}, but additionally counts how many
 * times it is evaluated (see {@link CountedNotBlankValidator}).
 * <p>
 * The count is what makes the {@code @Valid} cascade observable: in the workbook form the same row is handed to
 * the validator twice - once by the workbook-object pass that cascades through the sheet field, once by the
 * per-row/per-collection pass - and only a counting constraint can tell the two apart, because validation itself
 * is idempotent.
 */
@Target({FIELD})
@Retention(RUNTIME)
@Documented
@Constraint(validatedBy = {CountedNotBlankValidator.class})
public @interface CountedNotBlank {

    /**
     * The message template reported when the value is blank.
     *
     * @return the error message template used when the constraint is violated
     */
    String message() default "value must not be blank.";

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

}
