package io.github.hclimkr.pxl.internal.support;

import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.commons.lang3.time.DurationFormatUtils;

import java.time.Duration;
import java.time.Period;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Token-based parser that reads {@link Duration}/{@link Period} strings back from a {@link DurationFormatUtils}-style pattern.
 * <p>
 * On export, PXL formats {@link Duration}/{@link Period} values with a custom pattern via Apache Commons Lang's
 * {@code DurationFormatUtils.formatDuration}/{@code formatPeriod}. Commons Lang provides no reverse parser for that format
 * (it is format-only), so this class supplies the matching parser used on import: it compiles the pattern to a regex once
 * (via {@link #compileTemporalPattern(String)}) and extracts each time field from every value.
 */
public final class PxlTemporalAmountSupport {

    /**
     * Prevents instantiation.
     */
    private PxlTemporalAmountSupport() {

        throw new AssertionError("no instances of this class");
    }

    // Tokens used by DurationFormatUtils / DurationFormatUtils.formatPeriod. (y=year, M=month, d=day, H=hour, m=minute, s=second, S=millisecond)
    private static final String TEMPORAL_AMOUNT_PATTERN_TOKENS = "yMdHmsS";

    /**
     * Parses a string into a {@link Duration} using a pre-compiled {@link DurationFormatUtils}-style pattern (e.g. {@code "HH:mm:ss"}, {@code "d'd'H'h'm'm'"}).
     * <p>
     * Apache Commons Lang's {@link DurationFormatUtils} is format-only (no parser), so this provides a token-based parser to read back
     * strings that were emitted with the same pattern during export. Only the {@code d/H/m/s/S} tokens are meaningful for {@link Duration}; {@code y/M} tokens, if present, are ignored.
     * The pattern is a column-level constant, so it is compiled once (via {@link #compileTemporalPattern(String)}) and passed in here rather than recompiled per value.
     *
     * @param value           the string to parse
     * @param compiledPattern the compiled DurationFormatUtils-style pattern
     * @return the parsed Duration
     * @throws IllegalArgumentException when the value does not match the pattern or the value is outside the Duration range
     */
    public static Duration parseDurationByPattern(final String value,
                                                  final CompiledTemporalPattern compiledPattern) {

        final Map<Character, Long> fields = parseTemporalAmountFields(value, compiledPattern);

        try {
            return Duration.ZERO
                    .plusDays(fields.getOrDefault('d', 0L))
                    .plusHours(fields.getOrDefault('H', 0L))
                    .plusMinutes(fields.getOrDefault('m', 0L))
                    .plusSeconds(fields.getOrDefault('s', 0L))
                    .plusMillis(fields.getOrDefault('S', 0L));
        } catch (ArithmeticException e) {
            throw new IllegalArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_PERIOD_DURATION_DURATION_RANGE, value), e);
        }
    }

    /**
     * Parses a string into a {@link Period} using a pre-compiled {@link DurationFormatUtils}-style pattern.
     * <p>
     * Only the {@code y/M/d} tokens are meaningful for {@link Period}; {@code H/m/s/S} (time-component) tokens, if present, are ignored.
     * The pattern is a column-level constant, so it is compiled once (via {@link #compileTemporalPattern(String)}) and passed in here rather than recompiled per value.
     *
     * @param value           the string to parse
     * @param compiledPattern the compiled DurationFormatUtils-style pattern
     * @return the parsed Period
     * @throws IllegalArgumentException when the value does not match the pattern or the value is outside the int range
     */
    public static Period parsePeriodByPattern(final String value,
                                              final CompiledTemporalPattern compiledPattern) {

        final Map<Character, Long> fields = parseTemporalAmountFields(value, compiledPattern);

        try {
            return Period.of(
                    Math.toIntExact(fields.getOrDefault('y', 0L)),
                    Math.toIntExact(fields.getOrDefault('M', 0L)),
                    Math.toIntExact(fields.getOrDefault('d', 0L)));
        } catch (ArithmeticException e) {
            throw new IllegalArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_PERIOD_DURATION_INT_RANGE, value), e);
        }
    }

    /**
     * Compiles a {@link DurationFormatUtils}-style pattern into a reusable regular expression plus the capture-group field order.
     * <p>
     * A run of N identical tokens matches a digit group of at least N digits (N&gt;1) or one or more digits (N==1). Portions enclosed
     * in single quotes and all other characters are treated as literals. The produced regex is always well-formed (digit groups
     * and {@link Pattern#quote(String)} literals only), so this never throws for any pattern string. The result is meant to be
     * cached per column (the pattern is a column-level constant) and reused for every cell.
     *
     * @param pattern the DurationFormatUtils-style pattern
     * @return the compiled pattern (source pattern, regex, and field order)
     */
    public static CompiledTemporalPattern compileTemporalPattern(final String pattern) {

        final StringBuilder regex = new StringBuilder();
        final List<Character> fieldOrder = new ArrayList<>();

        int i = 0;
        while (i < pattern.length()) {
            final char c = pattern.charAt(i);

            if (c == '\'') {
                // Single-quote literal. ('' is a literal single quote)
                i++;
                final StringBuilder literal = new StringBuilder();
                if (i < pattern.length() && pattern.charAt(i) == '\'') {
                    literal.append('\'');
                    i++;
                } else {
                    while (i < pattern.length() && pattern.charAt(i) != '\'') {
                        literal.append(pattern.charAt(i));
                        i++;
                    }
                    i++; // consume the closing quote
                }
                regex.append(Pattern.quote(literal.toString()));
            } else if (TEMPORAL_AMOUNT_PATTERN_TOKENS.indexOf(c) >= 0) {
                int n = 0;
                while (i < pattern.length() && pattern.charAt(i) == c) {
                    n++;
                    i++;
                }
                regex.append(n == 1 ? "(\\d+)" : "(\\d{" + n + ",})");
                fieldOrder.add(c);
            } else {
                regex.append(Pattern.quote(String.valueOf(c)));
                i++;
            }
        }

        return new CompiledTemporalPattern(pattern, Pattern.compile(regex.toString()), Collections.unmodifiableList(fieldOrder));
    }

    /**
     * Matches a value against the compiled pattern and extracts each time field value.
     * <p>
     * Only the per-value {@link Matcher} and field extraction run here; the (constant) pattern was already compiled by {@link #compileTemporalPattern(String)}.
     *
     * @param value           the string to parse
     * @param compiledPattern the pre-compiled pattern to match the value against
     * @return a map from each time-field token (one of {@code yMdHmsS}) to its parsed numeric value
     * @throws IllegalArgumentException when the value does not match the pattern
     */
    private static Map<Character, Long> parseTemporalAmountFields(final String value,
                                                                  final CompiledTemporalPattern compiledPattern) {

        final Matcher matcher = compiledPattern.getRegex().matcher(value.trim());
        if (!matcher.matches()) {
            throw new IllegalArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_PERIOD_DURATION_PATTERN_MISMATCH, String.valueOf(value), compiledPattern.getSourcePattern()));
        }

        final List<Character> fieldOrder = compiledPattern.getFieldOrder();
        final Map<Character, Long> fields = new HashMap<>();
        for (int g = 0; g < fieldOrder.size(); g++) {
            fields.merge(fieldOrder.get(g), Long.parseLong(matcher.group(g + 1)), Long::sum);
        }

        return fields;
    }

    /**
     * A compiled DurationFormatUtils-style pattern: the source pattern string (for diagnostics), the derived regex, and the
     * capture-group field order (which time field each group holds). Immutable and thread-safe ({@link Pattern} is immutable;
     * a fresh {@link Matcher} is created per parse). Held on the column metadata so it is compiled once per column, not per cell.
     */
    @Getter
    @AllArgsConstructor
    public static final class CompiledTemporalPattern {

        /**
         * the original DurationFormatUtils-style pattern string (used for mismatch diagnostics).
         */
        private final String sourcePattern;

        /**
         * the compiled regular expression derived from the source pattern.
         */
        private final Pattern regex;

        /**
         * the field order: the time-field token each capture group (1-based) holds.
         */
        private final List<Character> fieldOrder;

    }

}
