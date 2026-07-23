package io.github.hclimkr.pxl.util;

import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * Null-safe collection and array helpers used throughout PXL: emptiness/size checks, safe indexed
 * access, and duplicate/uniqueness detection.
 */
public final class PxlCollectionUtils {

    /**
     * Prevents instantiation.
     */
    private PxlCollectionUtils() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Determines whether the collection is empty, treating {@code null} as empty.
     *
     * @param collection the collection to inspect, may be {@code null}
     * @return {@code true} if the collection is {@code null} or has no elements
     */
    public static boolean isEmpty(final Collection<?> collection) {

        return Objects.isNull(collection) || collection.isEmpty();
    }

    /**
     * Determines whether the collection contains at least one element, treating {@code null} as empty.
     *
     * @param collection the collection to inspect, may be {@code null}
     * @return {@code true} if the collection is non-{@code null} and has elements
     */
    public static boolean isNotEmpty(final Collection<?> collection) {

        return !isEmpty(collection);
    }

    /**
     * Returns the size of the collection, treating {@code null} as size 0.
     *
     * @param collection the collection to size, may be {@code null}
     * @return the element count, or 0 if {@code null}
     */
    public static int size(final Collection<?> collection) {

        return Objects.isNull(collection) ? 0 : collection.size();
    }

    /**
     * Returns the list element at the given index, or {@code null} for a {@code null} list or an
     * out-of-range index (no {@link IndexOutOfBoundsException} is thrown).
     *
     * @param list  the list to read from, may be {@code null}
     * @param index the zero-based index
     * @param <T>   the element type
     * @return the element, or {@code null} if the list is {@code null} or the index is out of range
     */
    public static <T> T get(final List<? extends T> list,
                            final int index) {

        if (Objects.isNull(list) || index < 0 || index >= size(list)) {
            return null;
        }

        return list.get(index);
    }

    /**
     * Returns the array element at the given index, or {@code null} for a {@code null} array or an
     * out-of-range index (no {@link ArrayIndexOutOfBoundsException} is thrown).
     *
     * @param array the array to read from, may be {@code null}
     * @param index the zero-based index
     * @param <T>   the element type
     * @return the element, or {@code null} if the array is {@code null} or the index is out of range
     */
    public static <T> T get(final T[] array,
                            final int index) {

        if (Objects.isNull(array) || index < 0 || index >= array.length) {
            return null;
        }

        return array[index];
    }

    /**
     * Returns the collection itself, or an immutable empty collection when it is {@code null}, so the
     * result is always safe to iterate.
     *
     * @param collection the collection, may be {@code null}
     * @param <T>        the element type
     * @return the original collection, or an empty collection if it was {@code null}
     */
    public static <T> Collection<T> emptyIfNull(final Collection<T> collection) {

        return Objects.isNull(collection) ? Collections.emptyList() : collection;
    }

    /**
     * Returns the set of elements that appear more than once in the collection. {@code null} elements
     * are ignored and a {@code null} collection yields an empty set.
     *
     * @param collection the collection to scan, may be {@code null}
     * @param <T>        the element type
     * @return the set of duplicated elements (empty if none)
     */
    public static <T> Set<T> findDuplicates(final Collection<? extends T> collection) {

        final Set<T> uniques = new HashSet<>();
        return emptyIfNull(collection).stream()
                .filter(e -> Objects.nonNull(e) && !uniques.add(e))   // Set.add() returns false if the element was already in the set.
                .collect(Collectors.toSet());
    }

    /**
     * Returns the set of mapped keys that occur more than once across the collection. Each element is
     * projected through {@code mapper}; {@code null} keys are ignored and a {@code null} collection
     * yields an empty set.
     *
     * @param collection the collection to scan, may be {@code null}
     * @param mapper     the function projecting each element to the key compared for duplication
     * @param <T>        the element type
     * @param <R>        the mapped key type
     * @return the set of duplicated keys (empty if none)
     */
    public static <T, R> Set<R> findDuplicates(final Collection<? extends T> collection,
                                               final Function<? super T, ? extends R> mapper) {

        final Set<R> uniques = new HashSet<>();
        return emptyIfNull(collection).stream()
                .map(mapper)
                .filter(e -> Objects.nonNull(e) && !uniques.add(e))   // Set.add() returns false if the element was already in the set.
                .collect(Collectors.toSet());
    }

    /**
     * Determines whether any element appears more than once, short-circuiting on the first duplicate.
     * {@code null} elements are ignored and a {@code null} collection has no duplicates.
     *
     * @param collection the collection to scan, may be {@code null}
     * @param <T>        the element type
     * @return {@code true} if a duplicate element exists
     */
    public static <T> boolean hasDuplicates(final Collection<? extends T> collection) {

        final Set<T> uniques = new HashSet<>();
        return emptyIfNull(collection).stream()
                .anyMatch(e -> Objects.nonNull(e) && !uniques.add(e));  // Set.add() returns false if the element was already in the set.
    }

    /**
     * Determines whether any mapped key occurs more than once, short-circuiting on the first duplicate.
     * {@code null} keys are ignored and a {@code null} collection has no duplicates.
     *
     * @param collection the collection to scan, may be {@code null}
     * @param mapper     the function projecting each element to the key compared for duplication
     * @param <T>        the element type
     * @param <R>        the mapped key type
     * @return {@code true} if a duplicate key exists
     */
    public static <T, R> boolean hasDuplicates(final Collection<? extends T> collection,
                                               final Function<? super T, ? extends R> mapper) {

        final Set<R> uniques = new HashSet<>();
        return emptyIfNull(collection).stream()
                .map(mapper)
                .anyMatch(e -> Objects.nonNull(e) && !uniques.add(e));  // Set.add() returns false if the element was already in the set.
    }

    /**
     * Determines whether every element is unique. A {@code null} collection is considered all-unique.
     *
     * @param collection the collection to check, may be {@code null}
     * @param <T>        the element type
     * @return {@code true} if the collection is {@code null} or contains no duplicates
     */
    public static <T> boolean hasAllUnique(final Collection<? extends T> collection) {

        return Objects.isNull(collection) || !hasDuplicates(collection);
    }

    /**
     * Determines whether the mapped keys of every element are unique. A {@code null} collection is
     * considered all-unique.
     *
     * @param collection the collection to check, may be {@code null}
     * @param mapper     the function projecting each element to the key compared for uniqueness
     * @param <T>        the element type
     * @param <R>        the mapped key type
     * @return {@code true} if the collection is {@code null} or its mapped keys contain no duplicates
     */
    public static <T, R> boolean hasAllUnique(final Collection<? extends T> collection,
                                              final Function<? super T, ? extends R> mapper) {

        return Objects.isNull(collection) || !hasDuplicates(collection, mapper);
    }

    /**
     * Determines whether all elements of the list are equal to each other (by {@link Objects#equals}).
     * An empty or {@code null} list is considered all-same.
     *
     * @param list the list to check, may be {@code null}
     * @param <T>  the element type
     * @return {@code true} if the list is empty/{@code null} or every element equals the first
     */
    public static <T> boolean hasAllSame(final List<T> list) {

        return isEmpty(list) || list.stream().allMatch(o -> Objects.equals(list.get(0), o));
    }

}
