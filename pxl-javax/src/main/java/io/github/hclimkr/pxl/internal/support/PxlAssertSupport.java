package io.github.hclimkr.pxl.internal.support;

import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;

import java.util.Collection;
import java.util.Map;
import java.util.Objects;
import java.util.function.Supplier;

/**
 * Argument-precondition utilities modeled after Apache Commons Lang3 {@link Validate}, but throwing
 * Pxl exceptions so that every failure stays within the {@link PxlException}
 * hierarchy.
 *
 * <p>A {@code null} required argument fails with {@link PxlNullPointerException}; an argument that is
 * present but otherwise invalid (empty, blank, or a false condition) fails with {@link PxlArgumentException}.
 * Both are checked exceptions, so callers declare {@code throws PxlNullPointerException} /
 * {@code throws PxlArgumentException} (or their common supertype {@link PxlException}).</p>
 *
 * <p>Each {@code notNull}/{@code notEmpty}/{@code notBlank} method comes in three forms: a no-argument form that
 * uses a generic default message; a form that takes the checked argument's <em>parameter name</em> (Java cannot
 * capture the caller's variable name automatically) and assembles a standard message from it - e.g.
 * {@code notNull(rowClass, "rowClass")} fails with {@code "argument 'rowClass' is null."} (the message is
 * localized; English is the default and Korean is available - see {@link PxlI18nDiagnostic}); and a form that takes a
 * {@link Supplier} of the exception to throw as-is on failure - e.g.
 * {@code notEmpty(tags, () -> new IllegalArgumentException("at least one tag is required"))} - letting the caller
 * raise a non-Pxl exception of their own choosing. {@code isTrue} validates a condition rather than a named argument,
 * so its second form takes a message.</p>
 *
 * <p>{@code notNegative} guards the index arguments of the public {@code util/} helpers. It comes in the named form
 * only, since an index failure that does not say which of several index arguments was rejected is of little use, and
 * it checks the lower bound alone - an upper bound is a property of the sheet's format rather than of the argument,
 * and POI rejects an index past its own limit with a message that names that limit.</p>
 */
public final class PxlAssertSupport {

    /**
     * Prevents instantiation.
     */
    private PxlAssertSupport() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Validates that the argument is not {@code null}, using a generic default message.
     *
     * @param object the object to check
     * @param <T>    the object type
     * @return the validated object (never {@code null})
     * @throws PxlNullPointerException if {@code object} is {@code null}
     */
    public static <T> T notNull(final T object)
            throws PxlNullPointerException {

        return notNull(object, (String) null);
    }

    /**
     * Validates that the argument is not {@code null}, building the message from the given parameter name.
     *
     * @param object        the object to check
     * @param parameterName the name of the checked parameter, embedded into the message
     * @param <T>           the object type
     * @return the validated object (never {@code null})
     * @throws PxlNullPointerException if {@code object} is {@code null}
     */
    public static <T> T notNull(final T object, final String parameterName)
            throws PxlNullPointerException {

        if (Objects.isNull(object)) {
            throw new PxlNullPointerException(nullMessage(parameterName));
        }

        return object;
    }

    /**
     * Validates that the argument is not {@code null}, throwing the supplied exception on failure.
     *
     * @param object            the object to check
     * @param exceptionSupplier supplies the exception thrown when the check fails
     * @param <T>               the object type
     * @param <X>               the thrown exception type
     * @return the validated object (never {@code null})
     * @throws X if {@code object} is {@code null}
     */
    public static <T, X extends Throwable> T notNull(final T object, final Supplier<? extends X> exceptionSupplier)
            throws X {

        if (Objects.isNull(object)) {
            throw exceptionSupplier.get();
        }

        return object;
    }

    /**
     * Validates that the array is neither {@code null} nor empty, using a generic default message.
     *
     * @param array the array to check
     * @param <T>   the array element type
     * @return the validated array (never {@code null} or empty)
     * @throws PxlNullPointerException if {@code array} is {@code null}
     * @throws PxlArgumentException    if {@code array} is empty
     */
    public static <T> T[] notEmpty(final T[] array)
            throws PxlNullPointerException, PxlArgumentException {

        return notEmpty(array, (String) null);
    }

    /**
     * Validates that the array is neither {@code null} nor empty, building the message from the given parameter name.
     *
     * @param array         the array to check
     * @param parameterName the name of the checked parameter, embedded into the message
     * @param <T>           the array element type
     * @return the validated array (never {@code null} or empty)
     * @throws PxlNullPointerException if {@code array} is {@code null}
     * @throws PxlArgumentException    if {@code array} is empty
     */
    public static <T> T[] notEmpty(final T[] array, final String parameterName)
            throws PxlNullPointerException, PxlArgumentException {

        if (Objects.isNull(array)) {
            throw new PxlNullPointerException(nullMessage(parameterName));
        }
        if (array.length == 0) {
            throw new PxlArgumentException(emptyMessage(parameterName));
        }

        return array;
    }

    /**
     * Validates that the array is neither {@code null} nor empty, throwing the supplied exception on failure.
     *
     * @param array             the array to check
     * @param exceptionSupplier supplies the exception thrown when the check fails
     * @param <T>               the array element type
     * @param <X>               the thrown exception type
     * @return the validated array (never {@code null} or empty)
     * @throws X if {@code array} is {@code null} or empty
     */
    public static <T, X extends Throwable> T[] notEmpty(final T[] array, final Supplier<? extends X> exceptionSupplier)
            throws X {

        if (Objects.isNull(array) || array.length == 0) {
            throw exceptionSupplier.get();
        }

        return array;
    }

    /**
     * Validates that the collection is neither {@code null} nor empty, using a generic default message.
     *
     * @param collection the collection to check
     * @param <T>        the collection type
     * @return the validated collection (never {@code null} or empty)
     * @throws PxlNullPointerException if {@code collection} is {@code null}
     * @throws PxlArgumentException    if {@code collection} is empty
     */
    public static <T extends Collection<?>> T notEmpty(final T collection)
            throws PxlNullPointerException, PxlArgumentException {

        return notEmpty(collection, (String) null);
    }

    /**
     * Validates that the collection is neither {@code null} nor empty, building the message from the given parameter name.
     *
     * @param collection    the collection to check
     * @param parameterName the name of the checked parameter, embedded into the message
     * @param <T>           the collection type
     * @return the validated collection (never {@code null} or empty)
     * @throws PxlNullPointerException if {@code collection} is {@code null}
     * @throws PxlArgumentException    if {@code collection} is empty
     */
    public static <T extends Collection<?>> T notEmpty(final T collection, final String parameterName)
            throws PxlNullPointerException, PxlArgumentException {

        if (Objects.isNull(collection)) {
            throw new PxlNullPointerException(nullMessage(parameterName));
        }
        if (collection.isEmpty()) {
            throw new PxlArgumentException(emptyMessage(parameterName));
        }

        return collection;
    }

    /**
     * Validates that the collection is neither {@code null} nor empty, throwing the supplied exception on failure.
     *
     * @param collection        the collection to check
     * @param exceptionSupplier supplies the exception thrown when the check fails
     * @param <T>               the collection type
     * @param <X>               the thrown exception type
     * @return the validated collection (never {@code null} or empty)
     * @throws X if {@code collection} is {@code null} or empty
     */
    public static <T extends Collection<?>, X extends Throwable> T notEmpty(final T collection, final Supplier<? extends X> exceptionSupplier)
            throws X {

        if (Objects.isNull(collection) || collection.isEmpty()) {
            throw exceptionSupplier.get();
        }

        return collection;
    }

    /**
     * Validates that the map is neither {@code null} nor empty, using a generic default message.
     *
     * @param map the map to check
     * @param <T> the map type
     * @return the validated map (never {@code null} or empty)
     * @throws PxlNullPointerException if {@code map} is {@code null}
     * @throws PxlArgumentException    if {@code map} is empty
     */
    public static <T extends Map<?, ?>> T notEmpty(final T map)
            throws PxlNullPointerException, PxlArgumentException {

        return notEmpty(map, (String) null);
    }

    /**
     * Validates that the map is neither {@code null} nor empty, building the message from the given parameter name.
     *
     * @param map           the map to check
     * @param parameterName the name of the checked parameter, embedded into the message
     * @param <T>           the map type
     * @return the validated map (never {@code null} or empty)
     * @throws PxlNullPointerException if {@code map} is {@code null}
     * @throws PxlArgumentException    if {@code map} is empty
     */
    public static <T extends Map<?, ?>> T notEmpty(final T map, final String parameterName)
            throws PxlNullPointerException, PxlArgumentException {

        if (Objects.isNull(map)) {
            throw new PxlNullPointerException(nullMessage(parameterName));
        }
        if (map.isEmpty()) {
            throw new PxlArgumentException(emptyMessage(parameterName));
        }

        return map;
    }

    /**
     * Validates that the map is neither {@code null} nor empty, throwing the supplied exception on failure.
     *
     * @param map               the map to check
     * @param exceptionSupplier supplies the exception thrown when the check fails
     * @param <T>               the map type
     * @param <X>               the thrown exception type
     * @return the validated map (never {@code null} or empty)
     * @throws X if {@code map} is {@code null} or empty
     */
    public static <T extends Map<?, ?>, X extends Throwable> T notEmpty(final T map, final Supplier<? extends X> exceptionSupplier)
            throws X {

        if (Objects.isNull(map) || map.isEmpty()) {
            throw exceptionSupplier.get();
        }

        return map;
    }

    /**
     * Validates that the character sequence is neither {@code null} nor empty (length 0), using a generic default message.
     *
     * @param chars the character sequence to check
     * @param <T>   the character-sequence type
     * @return the validated character sequence (never {@code null} or empty)
     * @throws PxlNullPointerException if {@code chars} is {@code null}
     * @throws PxlArgumentException    if {@code chars} is empty
     */
    public static <T extends CharSequence> T notEmpty(final T chars)
            throws PxlNullPointerException, PxlArgumentException {

        return notEmpty(chars, (String) null);
    }

    /**
     * Validates that the character sequence is neither {@code null} nor empty (length 0), building the message from the
     * given parameter name.
     *
     * @param chars         the character sequence to check
     * @param parameterName the name of the checked parameter, embedded into the message
     * @param <T>           the character-sequence type
     * @return the validated character sequence (never {@code null} or empty)
     * @throws PxlNullPointerException if {@code chars} is {@code null}
     * @throws PxlArgumentException    if {@code chars} is empty
     */
    public static <T extends CharSequence> T notEmpty(final T chars, final String parameterName)
            throws PxlNullPointerException, PxlArgumentException {

        if (Objects.isNull(chars)) {
            throw new PxlNullPointerException(nullMessage(parameterName));
        }
        if (chars.length() == 0) {
            throw new PxlArgumentException(emptyMessage(parameterName));
        }

        return chars;
    }

    /**
     * Validates that the character sequence is neither {@code null} nor empty (length 0), throwing the supplied exception on failure.
     *
     * @param chars             the character sequence to check
     * @param exceptionSupplier supplies the exception thrown when the check fails
     * @param <T>               the character-sequence type
     * @param <X>               the thrown exception type
     * @return the validated character sequence (never {@code null} or empty)
     * @throws X if {@code chars} is {@code null} or empty
     */
    public static <T extends CharSequence, X extends Throwable> T notEmpty(final T chars, final Supplier<? extends X> exceptionSupplier)
            throws X {

        if (StringUtils.isEmpty(chars)) {
            throw exceptionSupplier.get();
        }

        return chars;
    }

    /**
     * Validates that the character sequence is neither {@code null}, empty, nor whitespace-only, using a generic default message.
     *
     * @param chars the character sequence to check
     * @param <T>   the character-sequence type
     * @return the validated character sequence (never {@code null} or blank)
     * @throws PxlNullPointerException if {@code chars} is {@code null}
     * @throws PxlArgumentException    if {@code chars} is empty or whitespace-only
     */
    public static <T extends CharSequence> T notBlank(final T chars)
            throws PxlNullPointerException, PxlArgumentException {

        return notBlank(chars, (String) null);
    }

    /**
     * Validates that the character sequence is neither {@code null}, empty, nor whitespace-only, building the message
     * from the given parameter name.
     *
     * @param chars         the character sequence to check
     * @param parameterName the name of the checked parameter, embedded into the message
     * @param <T>           the character-sequence type
     * @return the validated character sequence (never {@code null} or blank)
     * @throws PxlNullPointerException if {@code chars} is {@code null}
     * @throws PxlArgumentException    if {@code chars} is empty or whitespace-only
     */
    public static <T extends CharSequence> T notBlank(final T chars, final String parameterName)
            throws PxlNullPointerException, PxlArgumentException {

        if (Objects.isNull(chars)) {
            throw new PxlNullPointerException(nullMessage(parameterName));
        }
        if (StringUtils.isBlank(chars)) {
            throw new PxlArgumentException(blankMessage(parameterName));
        }

        return chars;
    }

    /**
     * Validates that the character sequence is neither {@code null}, empty, nor whitespace-only, throwing the supplied
     * exception on failure.
     *
     * @param chars             the character sequence to check
     * @param exceptionSupplier supplies the exception thrown when the check fails
     * @param <T>               the character-sequence type
     * @param <X>               the thrown exception type
     * @return the validated character sequence (never {@code null} or blank)
     * @throws X if {@code chars} is {@code null}, empty, or whitespace-only
     */
    public static <T extends CharSequence, X extends Throwable> T notBlank(final T chars, final Supplier<? extends X> exceptionSupplier)
            throws X {

        if (StringUtils.isBlank(chars)) {
            throw exceptionSupplier.get();
        }

        return chars;
    }

    /**
     * Validates that the boolean expression is {@code true}, using a generic default message.
     *
     * @param expression the boolean expression to check
     * @throws PxlArgumentException if {@code expression} is {@code false}
     */
    public static void isTrue(final boolean expression)
            throws PxlArgumentException {

        isTrue(expression, PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.ASSERT_IS_TRUE));
    }

    /**
     * Validates that the boolean expression is {@code true}.
     *
     * @param expression the boolean expression to check
     * @param message    the exception message used when the check fails
     * @throws PxlArgumentException if {@code expression} is {@code false}
     */
    public static void isTrue(final boolean expression, final String message)
            throws PxlArgumentException {

        if (!expression) {
            throw new PxlArgumentException(message);
        }
    }

    /**
     * Validates that the boolean expression is {@code true}, throwing the supplied exception on failure.
     *
     * @param expression        the boolean expression to check
     * @param exceptionSupplier supplies the exception thrown when the check fails
     * @param <X>               the thrown exception type
     * @throws X if {@code expression} is {@code false}
     */
    public static <X extends Throwable> void isTrue(final boolean expression, final Supplier<? extends X> exceptionSupplier)
            throws X {

        if (!expression) {
            throw exceptionSupplier.get();
        }
    }

    /**
     * Validates that the index is not negative, building the message from the given parameter name.
     * <p>
     * Only the lower bound is checked. An upper bound belongs to the format the sheet is in rather than to the
     * argument, and POI already turns an index past its own limit down with a message that names the limit.
     *
     * @param index         the index to check
     * @param parameterName the name of the checked parameter, embedded into the message
     * @return the validated index (never negative)
     * @throws PxlArgumentException if {@code index} is negative
     */
    public static int notNegative(final int index, final String parameterName)
            throws PxlArgumentException {

        if (index < 0) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.ASSERT_NOT_NEGATIVE_NAMED, parameterName, String.valueOf(index)));
        }

        return index;
    }

    /**
     * Builds the localized "argument is null" message, using the named variant when a parameter name is given.
     *
     * @param parameterName the parameter name to embed, or blank/{@code null} for the generic message
     * @return the localized null-argument message
     */
    private static String nullMessage(final String parameterName) {

        return StringUtils.isBlank(parameterName)
                ? PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.ASSERT_NOT_NULL)
                : PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.ASSERT_NOT_NULL_NAMED, parameterName);
    }

    /**
     * Builds the localized "argument is empty" message, using the named variant when a parameter name is given.
     *
     * @param parameterName the parameter name to embed, or blank/{@code null} for the generic message
     * @return the localized empty-argument message
     */
    private static String emptyMessage(final String parameterName) {

        return StringUtils.isBlank(parameterName)
                ? PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.ASSERT_NOT_EMPTY)
                : PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.ASSERT_NOT_EMPTY_NAMED, parameterName);
    }

    /**
     * Builds the localized "argument is blank" message, using the named variant when a parameter name is given.
     *
     * @param parameterName the parameter name to embed, or blank/{@code null} for the generic message
     * @return the localized blank-argument message
     */
    private static String blankMessage(final String parameterName) {

        return StringUtils.isBlank(parameterName)
                ? PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.ASSERT_NOT_BLANK)
                : PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.ASSERT_NOT_BLANK_NAMED, parameterName);
    }

}
