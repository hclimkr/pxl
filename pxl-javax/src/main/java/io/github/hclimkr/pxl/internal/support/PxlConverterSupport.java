package io.github.hclimkr.pxl.internal.support;

import io.github.hclimkr.pxl.internal.constraint.Nullable;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.Objects;

/**
 * The reflective calls a column's custom converters are reached through, and the unwrapping their failures need.
 * <p>
 * A converter handle is resolved once when the column metadata is built - an export converter method, an import
 * converter method, a {@link String} constructor, a {@code toString} method - and the codecs then call it per value.
 * The calling itself is the same whatever the target type is, so it lives here; what differs between an enum column
 * and a custom object column is the fallback chain around the call and the diagnostic key a failure is reported
 * with, and both of those stay in the codec that owns the policy.
 * <p>
 * Nothing here builds an exception. A caller catches what these methods throw, runs the cause through
 * {@code unwrapInvocationCause}, and raises its own diagnostic - which is why the codecs keep their keys and this
 * class stays a plain invoker.
 */
public final class PxlConverterSupport {

    /**
     * Prevents instantiation.
     */
    private PxlConverterSupport() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Invokes an export converter method against the value being written out.
     * A static converter takes the value as its argument, an instance converter is called on the value itself.
     *
     * @param exportConverterMethod the resolved export converter method
     * @param value                 the value being converted
     * @return the string the converter produced, which may be {@code null} if the converter returns one
     * @throws ReflectiveOperationException if the method is inaccessible or throws while producing the string
     */
    public static String invokeExportConverter(final Method exportConverterMethod,
                                               final Object value)
            throws ReflectiveOperationException {

        if (Modifier.isStatic(exportConverterMethod.getModifiers())) {
            return (String) exportConverterMethod.invoke(null, value);
        }

        return (String) exportConverterMethod.invoke(value);
    }

    /**
     * Parses a string through the import converter method if there is one, and through the {@link String}
     * constructor otherwise. An import converter is always static, unlike its export counterpart.
     * <p>
     * Returns {@code null} when neither handle is available, leaving the caller to decide what that means: a custom
     * object column has nothing left to try, while an enum column still has its constants to scan.
     *
     * @param importConverterMethod the resolved import converter method; may be {@code null}
     * @param stringConstructor     the resolved {@link String} constructor; may be {@code null}
     * @param stringValue           the string being parsed
     * @return the parsed value, or {@code null} if neither handle is available
     * @throws ReflectiveOperationException if the handle is inaccessible, cannot be instantiated, or throws while parsing
     */
    public static Object invokeImportConverter(@Nullable final Method importConverterMethod,
                                               @Nullable final Constructor<?> stringConstructor,
                                               final String stringValue)
            throws ReflectiveOperationException {

        if (Objects.nonNull(importConverterMethod)) {
            return importConverterMethod.invoke(null, stringValue);
        }

        if (Objects.nonNull(stringConstructor)) {
            return stringConstructor.newInstance(stringValue);
        }

        return null;
    }

    /**
     * Returns what actually failed: reflection wraps whatever a converter throws in an
     * {@link InvocationTargetException}, and it is the wrapped cause that tells the user what went wrong.
     * Anything else is returned as it stands.
     *
     * @param throwable the throwable caught around a reflective call
     * @return the wrapped cause when the throwable is an {@link InvocationTargetException}; the throwable itself otherwise
     */
    public static Throwable unwrapInvocationCause(final Throwable throwable) {

        return (throwable instanceof InvocationTargetException) ? throwable.getCause() : throwable;
    }

}
