package io.github.hclimkr.pxl.internal.support;

import io.github.hclimkr.pxl.internal.constraint.Nullable;
import org.apache.commons.lang3.StringUtils;

import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * The name normalization and matching rules the import paths compare names with.
 * <p>
 * Import binds by name in three places - a sheet name, a column header, and an enum constant - and the three do not
 * share one rule. All three ignore whitespace, so the normalization is common; case is where they part. A sheet name
 * and an enum constant are matched ignoring case, while a column header is matched respecting it, and both halves of
 * that split are a documented contract rather than an accident. Keeping the rules here rather than at each call site
 * makes the difference visible in the method names: a caller that reaches for {@code matchesAnyRespectingCase} is
 * saying it must not fold case, not merely omitting to.
 * <p>
 * The comparisons take names that are already normalized, since a caller normalizes each name once and then matches
 * it against every candidate; {@code equalsNormalizedIgnoringCase} is the exception, normalizing the pair it is given
 * because its caller compares values it reads one at a time.
 */
public final class PxlNameMatchSupport {

    /**
     * Prevents instantiation.
     */
    private PxlNameMatchSupport() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Returns the name with all whitespace removed, which is the form every name comparison here expects.
     *
     * @param name the name to normalize; may be {@code null}
     * @return the name without whitespace, or {@code null} if {@code name} is {@code null}
     */
    public static String normalizeName(@Nullable final String name) {

        return StringUtils.deleteWhitespace(name);
    }

    /**
     * Returns the names normalized with {@code normalizeName}, dropping the entries that are blank once normalized.
     * The order of the surviving names is kept, since candidate names are matched in the order they were declared.
     *
     * @param names the names to normalize; may be {@code null}
     * @return a new list of the normalized, non-blank names; empty if {@code names} is {@code null} or holds none
     */
    public static List<String> normalizeNames(@Nullable final List<String> names) {

        if (Objects.isNull(names)) {
            return Collections.emptyList();
        }

        return names.stream()
                .map(PxlNameMatchSupport::normalizeName)
                .filter(StringUtils::isNotBlank)
                .collect(Collectors.toList());
    }

    /**
     * Returns whether any of the candidate names equals the given name, ignoring case.
     * Both sides are expected to be normalized already.
     *
     * @param candidateNames the candidate names to match against
     * @param name           the name read from the source; may be {@code null}
     * @return {@code true} if any candidate equals {@code name} ignoring case; {@code false} otherwise
     */
    public static boolean matchesAnyIgnoringCase(final List<String> candidateNames,
                                                 @Nullable final String name) {

        return candidateNames.stream()
                .anyMatch(candidateName -> StringUtils.equalsIgnoreCase(candidateName, name));
    }

    /**
     * Returns whether any of the candidate names equals the given name, respecting case.
     * Both sides are expected to be normalized already.
     *
     * @param candidateNames the candidate names to match against
     * @param name           the name read from the source; may be {@code null}
     * @return {@code true} if any candidate equals {@code name} exactly; {@code false} otherwise
     */
    public static boolean matchesAnyRespectingCase(final List<String> candidateNames,
                                                   @Nullable final String name) {

        return candidateNames.stream()
                .anyMatch(candidateName -> StringUtils.equals(candidateName, name));
    }

    /**
     * Returns whether the two names are equal once normalized, ignoring case; the same reference (including both
     * {@code null}) is equal, while a single {@code null} is not.
     * Unlike the matching methods above, this normalizes the pair it is given.
     *
     * @param name1 the first name; may be {@code null}
     * @param name2 the second name; may be {@code null}
     * @return {@code true} if the two names are considered equal
     */
    public static boolean equalsNormalizedIgnoringCase(@Nullable final String name1,
                                                       @Nullable final String name2) {

        if (name1 == name2) {
            return true;
        }

        if (Objects.isNull(name1) || Objects.isNull(name2)) {
            return false;
        }

        return StringUtils.equalsIgnoreCase(normalizeName(name1), normalizeName(name2));
    }

}
