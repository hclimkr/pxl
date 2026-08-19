package io.github.hclimkr.pxl.tcdata;

import javax.validation.ConstraintValidator;
import javax.validation.ConstraintValidatorContext;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Validator behind {@link CountedNotBlank}. Counts every evaluation in a static counter so a test can assert how
 * many times bean validation visited the same row.
 * <p>
 * Bean Validation instantiates this class reflectively and may reuse the instance, hence the static counter.
 * A test resets it with {@code resetEvaluations()} right before the operation under test.
 */
public class CountedNotBlankValidator implements ConstraintValidator<CountedNotBlank, CharSequence> {

    private static final AtomicInteger EVALUATIONS = new AtomicInteger();

    /**
     * Clears the evaluation counter.
     */
    public static void resetEvaluations() {

        EVALUATIONS.set(0);
    }

    /**
     * Returns how many times the constraint was evaluated since the last reset.
     *
     * @return the evaluation count
     */
    public static int evaluations() {

        return EVALUATIONS.get();
    }

    @Override
    public boolean isValid(final CharSequence value, final ConstraintValidatorContext context) {

        EVALUATIONS.incrementAndGet();

        return value != null && value.toString().trim().length() > 0;
    }

}
