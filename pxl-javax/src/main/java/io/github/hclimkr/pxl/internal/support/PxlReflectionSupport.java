package io.github.hclimkr.pxl.internal.support;

import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlReflectionException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import org.apache.commons.lang3.ClassUtils;

import java.lang.annotation.Annotation;
import java.lang.reflect.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/**
 * Reflection helpers for the binder: no-arg instantiation, field/method access, annotation member lookup, and
 * generic element-type resolution.
 */
public final class PxlReflectionSupport {

    /**
     * Prevents instantiation.
     */
    private PxlReflectionSupport() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Instantiates the class via its no-argument constructor, making it accessible even if non-public.
     *
     * @param clazz the class to instantiate
     * @return a new instance of the class
     * @throws PxlReflectionException when no no-arg constructor exists, the class cannot be instantiated (abstract/inaccessible),
     *                                or the constructor itself throws (the cause is propagated)
     */
    public static Object newClassInstance(final Class<?> clazz)
            throws PxlReflectionException {

        try {
            final Constructor<?> constructor = clazz.getDeclaredConstructor();
            constructor.setAccessible(true);
            return constructor.newInstance();
        } catch (InvocationTargetException e) {
            // Propagate the exception thrown inside the constructor
            throw new PxlReflectionException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_REFLECT_INSTANTIATE_ERROR, clazz.getSimpleName()), e.getCause());
        } catch (ReflectiveOperationException e) {
            // NoSuchMethodException (no no-arg constructor) / InstantiationException (abstract) / IllegalAccessException
            throw new PxlReflectionException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_REFLECT_NO_NO_ARG_CONSTRUCTOR, clazz.getSimpleName()), e);
        }
    }

    /**
     * Reads the value of the given field from the given object, making the field accessible first.
     *
     * @param field  the field to read
     * @param object the instance to read from
     * @return the field's value
     * @throws PxlReflectionException when the field cannot be accessed
     */
    public static Object getFieldValue(final Field field,
                                       final Object object)
            throws PxlReflectionException {

        try {
            field.setAccessible(true);
            return field.get(object);
        } catch (Exception e) {
            throw new PxlReflectionException(e);
        }
    }

    /**
     * Writes the given value into the given field of the given object, making the field accessible first.
     *
     * @param field  the field to write
     * @param object the instance to write to
     * @param value  the value to set
     * @throws PxlReflectionException when the field cannot be accessed or the value type is incompatible
     */
    public static void setFieldValue(final Field field,
                                     final Object object,
                                     final Object value)
            throws PxlReflectionException {

        try {
            field.setAccessible(true);
            field.set(object, value);
        } catch (Exception e) {
            throw new PxlReflectionException(e);
        }
    }

    /**
     * Collects and returns the fields of a class across the entire inheritance chain.
     * <p>
     * Fields declared in the subclass (derived) class are listed first, followed by fields of the superclass.
     * (The {@code override*} options rely on this order so that derived fields are processed first and claim a matching column name.)
     * Since the inheritance chain is collected as-is, a shadowed field name may be included more than once.
     *
     * @param clazz the class whose inheritance chain is scanned
     * @return the fields from the class and all superclasses, subclass fields first
     */
    public static List<Field> getAllFields(final Class<?> clazz) {

        final List<Field> fields = new ArrayList<>();
        for (Class<?> c = clazz; Objects.nonNull(c); c = c.getSuperclass()) {
//            fields.addAll(0, Arrays.asList(c.getDeclaredFields()));
            fields.addAll(Arrays.asList(c.getDeclaredFields()));
        }
        return fields;
    }

    /**
     * Collects and returns the methods of a class across the entire inheritance chain.
     * <p>
     * Methods declared in the subclass (derived) class are listed first, followed by methods of the superclass.
     * Since the inheritance chain is collected as-is, an overridden method may be included more than once (derived and super).
     *
     * @param clazz the class whose inheritance chain is scanned
     * @return the methods from the class and all superclasses, subclass methods first
     */
    public static List<Method> getAllMethods(final Class<?> clazz) {

        final List<Method> methods = new ArrayList<>();
        for (Class<?> c = clazz; Objects.nonNull(c); c = c.getSuperclass()) {
//            methods.addAll(0, Arrays.asList(c.getDeclaredMethods()));
            methods.addAll(Arrays.asList(c.getDeclaredMethods()));
        }
        return methods;
    }

    /**
     * Returns the first method (across the inheritance chain) that carries the specified annotation, made accessible.
     *
     * @param clazz           the class whose inheritance chain is scanned
     * @param annotationClazz the annotation type to look for
     * @return the first annotated method, or {@code null} if none is found
     */
    public static Method getAnnotatedMethod(final Class<?> clazz,
                                            final Class<? extends Annotation> annotationClazz) {

        final List<Method> methods = getAllMethods(clazz);
        for (final Method method : methods) {
            final Annotation annotation = method.getAnnotation(annotationClazz);
            if (Objects.nonNull(annotation)) {
                method.setAccessible(true);
                return method;
            }
        }

        return null;
    }

    /**
     * Changes the annotation value for the given key of the given annotation to newValue and
     * returns the previous value.
     */
    /* Disabled because it requires --add-opens java.base/sun.reflect.annotation=ALL-UNNAMED.
    @Deprecated
    @SuppressWarnings("unchecked")
    public static Object changeAnnotationValue(final Annotation annotation,
                                               final String key,
                                               final Object newValue) {

        final Object handler = Proxy.getInvocationHandler(annotation);

        Field memberValuesField;
        try {
            memberValuesField = handler.getClass().getDeclaredField("memberValues");
        } catch (NoSuchFieldException | SecurityException e) {
            throw new IllegalStateException(e);
        }
        memberValuesField.setAccessible(true);

        Map<String, Object> memberValues;
        try {
            memberValues = (Map<String, Object>) memberValuesField.get(handler);
        } catch (IllegalArgumentException | IllegalAccessException e) {
            throw new IllegalStateException(e);
        }

        Object oldValue = memberValues.get(key);
        if (Objects.nonNull(oldValue)
                && Objects.nonNull(newValue)
                && !Objects.equals(oldValue.getClass(), newValue.getClass())) {
            throw new IllegalArgumentException();
        }
        memberValues.put(key, newValue);

        return oldValue;
    }
    */

    /**
     * Creates a new annotation instance in which only the value of a single member (key) of the given annotation is changed to newValue.
     * Without mutating the original, returns a dynamic proxy that implements the same annotation type.
     * Since it does not rely on JDK-internal APIs, it works on both Java 8 (javax) and Java 17 (jakarta).
     *
     * @param annotation the source annotation whose members are delegated to
     * @param key        the name of the single member to override
     * @param newValue   the new value for that member
     * @param <T>        the annotation type
     * @return a proxy of the same annotation type returning {@code newValue} for {@code key} and the original values otherwise
     * @throws PxlReflectionException when the annotation type has no member named {@code key}
     * @throws PxlArgumentException   when {@code newValue} is {@code null} or not assignable to the member's type
     */
    @SuppressWarnings("unchecked")
    public static <T extends Annotation> T withAnnotationValue(final T annotation,
                                                               final String key,
                                                               final Object newValue)
            throws PxlReflectionException, PxlArgumentException {

        final Class<? extends Annotation> annotationType = annotation.annotationType();

        // 1) Verify the member exists
        final Method member;
        try {
            member = annotationType.getDeclaredMethod(key);
        } catch (NoSuchMethodException e) {
            throw new PxlReflectionException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_REFLECT_MEMBER_MISSING, annotationType.getSimpleName(), key), e);
        }

        // 2) Verify the value type (primitive members are compared as their wrappers; annotation members cannot be null)
        final Class<?> expectedType = ClassUtils.primitiveToWrapper(member.getReturnType());
        if (Objects.isNull(newValue) || !expectedType.isInstance(newValue)) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_REFLECT_MEMBER_INCOMPATIBLE, key, member.getReturnType().getSimpleName()));
        }

        // 3) Return newValue only for the key member; delegate the rest (other members + annotationType/equals/hashCode/toString) to the original
        final InvocationHandler invocationHandler = (proxy, method, args) -> {
            if (method.getParameterCount() == 0 && method.getName().equals(key)) {
                return newValue;
            }
            return method.invoke(annotation, args);
        };

        return (T) Proxy.newProxyInstance(
                annotationType.getClassLoader(),
                new Class<?>[]{annotationType},
                invocationHandler);
    }

    /**
     * Tests whether the class has a public constructor taking a single {@link String} argument.
     *
     * @param clazz the class to inspect
     * @return {@code true} if a {@code (String)} constructor exists
     */
    public static boolean hasStringTypeConstructor(final Class<?> clazz) {

        try {
            final Constructor<?> constructor = clazz.getConstructor(String.class);
            return true;
        } catch (NoSuchMethodException e) {
            return false;
        }
    }

    /**
     * Returns the class's public {@code (String)} constructor, made accessible.
     *
     * @param clazz the class to inspect
     * @return the {@code (String)} constructor, or {@code null} if none exists
     */
    public static Constructor<?> getStringTypeConstructor(final Class<?> clazz) {

        Constructor<?> constructor = null;

        try {
            constructor = clazz.getConstructor(String.class);
            constructor.setAccessible(true);
        } catch (NoSuchMethodException ignored) {
        }

        return constructor;
    }

//    public static boolean hasToStringMethod(final Class<?> clazz) {
//
//        try {
//            getToStringMethod(clazz);
//            return true;
//        } catch (NoSuchMethodException e) {
//            return false;
//        }
//    }

    /**
     * Returns the {@code toString()} method that the class (or a superclass) explicitly overrides, excluding
     * {@link Object#toString()}. Walks the inheritance chain up to but not including {@link Object}.
     *
     * @param clazz the class to inspect
     * @return the overriding {@code toString()} method, or {@code null} if only {@link Object}'s default exists
     */
    public static Method getToStringMethod(final Class<?> clazz) {

        // Excluding Object's default toString(), find a toString() that clazz or one of its superclasses explicitly overrides.
        // (getDeclaredMethod finds only directly declared methods, so walk the inheritance chain manually.)
        for (Class<?> c = clazz; Objects.nonNull(c) && c != Object.class; c = c.getSuperclass()) {
            try {
                final Method toStringMethod = c.getDeclaredMethod("toString");
                if (Objects.nonNull(toStringMethod)
                        && toStringMethod.getReturnType() == String.class
                        && toStringMethod.getParameterCount() == 0) {
                    return toStringMethod;
                }
            } catch (NoSuchMethodException ignored) {
                // not on this class -> keep searching up the superclass chain.
            }
        }

        return null;
    }

    /**
     * Returns the first generic type argument of a parameterized field (e.g. the element type {@code E} of {@code List<E>}).
     * <p>
     * Requires the field's generic type to be a {@link ParameterizedType} whose first argument is a concrete {@link Class};
     * raw types, wildcards, and nested generics are rejected.
     *
     * @param field the field to inspect
     * @return the first type argument as a concrete class
     * @throws PxlReflectionException when the field is a raw type or its first type argument is not a concrete class
     */
    public static Class<?> getParameterizedArgument0(final Field field)
            throws PxlReflectionException {

        final Type genericType = field.getGenericType();
        if (!(genericType instanceof ParameterizedType)) {
            throw new PxlReflectionException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_REFLECT_RAW_TYPE, field.getName()));
        }

        final ParameterizedType parameterizedType = (ParameterizedType) genericType;
        final Type argument0Type = parameterizedType.getActualTypeArguments()[0];
        if (!(argument0Type instanceof Class)) {
            throw new PxlReflectionException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_REFLECT_ELEMENT_TYPE_UNSUPPORTED, field.getName(), argument0Type.getTypeName()));
        }

        return (Class<?>) argument0Type;
    }

}
