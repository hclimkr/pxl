package io.github.hclimkr.pxl.internal.constraint;

import io.github.hclimkr.pxl.constraint.PxlByteSize;

import javax.validation.ConstraintValidator;
import javax.validation.ConstraintValidatorContext;

import java.nio.charset.Charset;

/**
 * Check that the byte length of a character sequence is between min and max.
 */
public class PxlByteSizeValidator implements ConstraintValidator<PxlByteSize, CharSequence> {

    /**
     * Byte length the element must be higher or equal to
     */
    private int min;

    /**
     * Byte length the element must be lower or equal to
     */
    private int max;

    /**
     * The charset used in parse to a string.
     */
    private Charset charset;

    /**
     * Initialize validator.
     *
     * @param constraintAnnotation annotation instance for a given constraint declaration
     * @throws IllegalArgumentException failed to get a charset by name, or min and max are invalid.
     */
    @Override
    public void initialize(PxlByteSize constraintAnnotation) {

        charset = Charset.forName(constraintAnnotation.charset());

        min = constraintAnnotation.min();
        max = constraintAnnotation.max();

        if (min < 0) {
            throw new IllegalArgumentException("min[" + min + "] must not be negative value.");
        }
        if (max < 0) {
            throw new IllegalArgumentException("max[" + max + "] must not be negative value.");
        }
        if (max < min) {
            throw new IllegalArgumentException("max[" + max + "] must be higher or equal to min[" + min + "].");
        }
    }

    /**
     * Validate execute.
     *
     * @param value   object to validate
     * @param context context in which the constraint is evaluated
     * @return {@code true} if {@code value} byte length is between the specified minimum and maximum, or null. otherwise {@code false}.
     */
    @Override
    public boolean isValid(CharSequence value, ConstraintValidatorContext context) {

        if (value == null) {
            return true;
        }

        int byteLength = value.toString().getBytes(charset).length;
        return min <= byteLength && byteLength <= max;
    }

}
