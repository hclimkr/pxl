package io.github.hclimkr.pxl.internal.support;

import io.github.hclimkr.pxl.exception.PxlReflectionException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.*;
import java.util.*;

/**
 * Class/type predicates that drive codec dispatch.
 * <p>
 * Provides the {@code is*Class} checks the resolver uses to classify a field's declared type (number, string, boolean,
 * character, date/time, UUID, collection, enum, custom object), plus resolution of a concrete {@link Collection}
 * implementation for a requested collection interface.
 */
public final class PxlClassSupport {

    /**
     * Prevents instantiation.
     */
    private PxlClassSupport() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Resolves a concrete, instantiable {@link Collection} class for the requested collection type.
     * <p>
     * When {@code collectionClass} is an interface, a default implementation is chosen: {@link ArrayList} for
     * {@link List}/{@link Collection}, {@link TreeSet} for {@link SortedSet}, {@link HashSet} for {@link Set},
     * {@link ArrayDeque} for {@link Deque}, and {@link LinkedList} for {@link Queue}. When it is a concrete class,
     * it is returned as-is provided it is a {@link Collection}. The chosen class is verified to be assignable to the
     * requested type.
     *
     * @param collectionClass the requested collection type (interface or concrete class)
     * @return a concrete collection class assignable to {@code collectionClass}
     * @throws PxlReflectionException when the type is not a supported collection or no compatible implementation exists
     */
    public static Class<? extends Collection<?>> getConcreteCollectionClass(final Class<?> collectionClass)
            throws PxlReflectionException {

        Class<?> concreteCollectionClass;

        if (collectionClass.isInterface()) {
            if (List.class.isAssignableFrom(collectionClass)) {
                concreteCollectionClass = ArrayList.class;
            } else if (SortedSet.class.isAssignableFrom(collectionClass)) {
                concreteCollectionClass = TreeSet.class;
            } else if (Set.class.isAssignableFrom(collectionClass)) {
                concreteCollectionClass = HashSet.class;
            } else if (Deque.class.isAssignableFrom(collectionClass)) {
                concreteCollectionClass = ArrayDeque.class;
            } else if (Queue.class.isAssignableFrom(collectionClass)) {
                concreteCollectionClass = LinkedList.class;
            } else if (Collection.class.isAssignableFrom(collectionClass)) {
                concreteCollectionClass = ArrayList.class;
            } else {
                throw new PxlReflectionException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_TYPE_UNSUPPORTED, collectionClass.getSimpleName()));
            }
        } else {
            if (Collection.class.isAssignableFrom(collectionClass)) {
                concreteCollectionClass = collectionClass;
            } else {
                throw new PxlReflectionException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_TYPE_UNSUPPORTED, collectionClass.getSimpleName()));
            }
        }

        if (!collectionClass.isAssignableFrom(concreteCollectionClass)) {
            throw new PxlReflectionException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_TYPE_UNSUPPORTED, collectionClass.getSimpleName()));
        }

        @SuppressWarnings("unchecked") final Class<? extends Collection<?>> result =
                (Class<? extends Collection<?>>) concreteCollectionClass;
        return result;
    }

    /**
     * Tests whether the type is a supported numeric type: the primitive and wrapper forms of byte/short/int/long/double/float,
     * or {@link BigInteger}/{@link BigDecimal}.
     *
     * @param clazz the type to test
     * @return {@code true} if the type is a supported numeric type
     */
    public static boolean isNumberClass(final Class<?> clazz) {

        if (clazz == byte.class) {
            return true;
        } else if (clazz == Byte.class) {
            return true;
        } else if (clazz == short.class) {
            return true;
        } else if (clazz == Short.class) {
            return true;
        } else if (clazz == int.class) {
            return true;
        } else if (clazz == Integer.class) {
            return true;
        } else if (clazz == long.class) {
            return true;
        } else if (clazz == Long.class) {
            return true;
        } else if (clazz == double.class) {
            return true;
        } else if (clazz == Double.class) {
            return true;
        } else if (clazz == float.class) {
            return true;
        } else if (clazz == Float.class) {
            return true;
        } else if (clazz == BigInteger.class) {
            return true;
        } else if (clazz == BigDecimal.class) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * Tests whether the type is {@link String}.
     *
     * @param clazz the type to test
     * @return {@code true} if the type is {@link String}
     */
    public static boolean isStringClass(final Class<?> clazz) {

        if (clazz == String.class) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * Tests whether the type is {@code boolean} or {@link Boolean}.
     *
     * @param clazz the type to test
     * @return {@code true} if the type is a boolean type
     */
    public static boolean isBooleanClass(final Class<?> clazz) {

        if (clazz == boolean.class) {
            return true;
        } else if (clazz == Boolean.class) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * Tests whether the type is {@code char} or {@link Character}.
     *
     * @param clazz the type to test
     * @return {@code true} if the type is a character type
     */
    public static boolean isCharacterClass(final Class<?> clazz) {

        if (clazz == char.class) {
            return true;
        } else if (clazz == Character.class) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * Tests whether the type is the legacy {@link Date}.
     *
     * @param clazz the type to test
     * @return {@code true} if the type is {@link Date}
     */
    public static boolean isJavaDateClass(final Class<?> clazz) {

        if (clazz == Date.class) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * Tests whether the type is one of the {@code java.time} temporal types:
     * {@link LocalDate}, {@link LocalTime}, {@link LocalDateTime}, {@link ZonedDateTime}, {@link OffsetTime}, or {@link OffsetDateTime}.
     *
     * @param clazz the type to test
     * @return {@code true} if the type is a supported {@code java.time} date/time type
     */
    public static boolean isDateTimeClass(final Class<?> clazz) {

        if (clazz == LocalDate.class) {
            return true;
        } else if (clazz == LocalTime.class) {
            return true;
        } else if (clazz == LocalDateTime.class) {
            return true;
        } else if (clazz == ZonedDateTime.class) {
            return true;
        } else if (clazz == OffsetTime.class) {
            return true;
        } else if (clazz == OffsetDateTime.class) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * Tests whether the type is one of the {@code java.time} temporal-amount types: {@link Period} or {@link Duration}.
     *
     * @param clazz the type to test
     * @return {@code true} if the type is {@link Period} or {@link Duration}
     */
    public static boolean isTemporalAmountClass(final Class<?> clazz) {

        if (clazz == Period.class) {
            return true;
        } else if (clazz == Duration.class) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * Tests whether the type is {@link UUID}.
     *
     * @param clazz the type to test
     * @return {@code true} if the type is {@link UUID}
     */
    public static boolean isUuidClass(final Class<?> clazz) {

        if (clazz == UUID.class) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * Tests whether the type is (or implements) {@link Collection}.
     *
     * @param clazz the type to test
     * @return {@code true} if the type is assignable to {@link Collection}
     */
    public static boolean isCollectionClass(final Class<?> clazz) {

        return Collection.class.isAssignableFrom(clazz);
    }

    /**
     * Tests whether the type is directly supported as a column value type by a built-in codec.
     * <p>
     * Covers the numeric, string, boolean, character, legacy {@link Date}, and {@code java.time} date/time types,
     * plus {@link Duration}, {@link Period}, {@link UUID}, any enum, and any {@link Collection}.
     *
     * @param clazz the type to test
     * @return {@code true} if a built-in codec handles the type
     */
    public static boolean isSupportedColumnClass(final Class<?> clazz) {

        return isNumberClass(clazz)
                || isStringClass(clazz)
                || isBooleanClass(clazz)
                || isCharacterClass(clazz)
                || isJavaDateClass(clazz)
                || isDateTimeClass(clazz)
                || isTemporalAmountClass(clazz)
                || isUuidClass(clazz)
                || clazz.isEnum()
                || isCollectionClass(clazz);
    }

    /**
     * Tests whether the type is a custom (non-built-in) type, i.e. the complement of {@link #isSupportedColumnClass(Class)}.
     * Such types are handled by the object codec via a String constructor / {@code toString} or a custom converter.
     *
     * @param clazz the type to test
     * @return {@code true} if no built-in codec directly supports the type
     */
    public static boolean isCustomClass(final Class<?> clazz) {

        return !isSupportedColumnClass(clazz);
    }

    /**
     * Tests whether the type is eligible for custom String conversion: any enum, or any custom (non-built-in) type.
     *
     * @param clazz the type to test
     * @return {@code true} if the type is an enum or a custom type
     */
    public static boolean isCustomConvertableClass(final Class<?> clazz) {

        return clazz.isEnum() || PxlClassSupport.isCustomClass(clazz);
    }

}
