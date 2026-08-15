package io.github.hclimkr.pxl.internal.support;

import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.ParsePosition;
import java.util.Locale;
import java.util.Objects;

/**
 * Numeric helpers shared by the number codecs: pattern formatting and parsing, float widening, and the
 * range/finiteness guards that keep a spreadsheet's {@code double} storage from silently corrupting a bound value.
 * <p>
 * A spreadsheet holds every number as a {@code double}, which is where the guards come in. {@code requireWithinRange}
 * rejects a value that would not survive the target type (including the 2^53 limit beyond which a {@code double}
 * can no longer represent consecutive integers), and the {@code requireFinite*} methods reject {@code NaN} and
 * infinity on both directions, since neither has a cell representation. All of them fail loudly rather than
 * truncating.
 * <p>
 * {@code getDecimalFormat} builds formatters on {@link Locale#ROOT} symbols so that a pattern such as
 * {@code "#,##0.##"} parses and formats identically regardless of the JVM's default locale, and
 * {@code floatToPlainDouble} widens a {@code float} through its short decimal form so {@code 0.1f} becomes
 * {@code 0.1} rather than {@code 0.10000000149011612}.
 * <p>
 * {@code parseFullyAsNumber} is what the codecs parse a patterned string with. It exists because {@code DecimalFormat}'s
 * own {@code parse(String)} stops at the first character the pattern cannot read and returns what it got so far,
 * so a value such as {@code "1e3"} would bind as {@code 1} without any error - the opposite of the pattern-less
 * path, where {@code Integer.parseInt} rejects the whole string.
 */
public final class PxlNumberSupport {

    /**
     * Prevents instantiation.
     */
    private PxlNumberSupport() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Creates a {@link DecimalFormat} for the given number pattern with locale-independent symbols
     * ({@link Locale#ROOT}): the decimal separator is {@code '.'} and the grouping separator is {@code ','},
     * regardless of the JVM default locale. This keeps a pattern such as {@code "#,##0.##"} parsing and
     * formatting identically on every machine (e.g. {@code de_DE}/{@code fr_FR}, where the default symbols
     * would otherwise swap {@code '.'} and {@code ','}).
     *
     * @param pattern the {@link DecimalFormat} pattern
     * @return a formatter using {@code Locale.ROOT} symbols
     * @throws IllegalArgumentException if the pattern is invalid
     */
    public static DecimalFormat getDecimalFormat(final String pattern) {

        return new DecimalFormat(pattern, DecimalFormatSymbols.getInstance(Locale.ROOT));
    }

    /**
     * Parses the whole string with the given import formatter, rejecting anything the pattern does not consume.
     * <p>
     * {@code DecimalFormat.parse(String)} only fails when nothing at all could be read: it parses as far as the pattern
     * matches, leaves the rest of the string behind and reports success, so {@code "123abc"} yields {@code 123} and
     * {@code "1e3"} yields {@code 1}. Parsing through a {@link ParsePosition} instead makes it possible to require that
     * the position ends at the end of the string, which is what the pattern-less paths ({@link Integer#parseInt(String)}
     * and friends) and the {@code java.time} codecs already do.
     * <p>
     * Note that this checks consumption only. A grouping separator in an unexpected place ({@code "1,2,3"}) is still
     * accepted, because {@code DecimalFormat} does not verify group sizes while parsing.
     *
     * @param formatter the column's import formatter
     * @param value     the string to parse
     * @param typeName  the target type name, used in the error message
     * @return the parsed number
     * @throws PxlCellCodecException when the string is not a number in this pattern or is only partly consumed by it
     */
    public static Number parseFullyAsNumber(final DecimalFormat formatter,
                                            final String value,
                                            final String typeName)
            throws PxlCellCodecException {

        final ParsePosition parsePosition = new ParsePosition(0);
        final Number number = formatter.parse(value, parsePosition);

        if (Objects.isNull(number) || parsePosition.getIndex() != value.length()) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(value), typeName));
        }

        return number;
    }

    /**
     * Parses the whole string with the given import formatter and returns it as a {@link BigDecimal}, for the unbounded
     * target types ({@link BigDecimal}/{@link BigInteger}) whose formatters run with {@code setParseBigDecimal(true)}.
     * <p>
     * That flag does not cover every result: infinity and NaN come back as a {@link Double} even with it set, so a plain
     * cast to {@link BigDecimal} would fail with a {@link ClassCastException} whose message says nothing about the cell.
     * Neither value has a {@link BigDecimal} form, so they are rejected here with the same message the {@link Double} and
     * {@link Float} codecs use.
     *
     * @param formatter the column's import formatter
     * @param value     the string to parse
     * @param typeName  the target type name, used in the error message
     * @return the parsed value as a {@link BigDecimal}
     * @throws PxlCellCodecException when the string is not a number in this pattern, is only partly consumed by it, or is
     *                               infinity/NaN
     */
    public static BigDecimal parseFullyAsBigDecimal(final DecimalFormat formatter,
                                                    final String value,
                                                    final String typeName)
            throws PxlCellCodecException {

        final Number number = parseFullyAsNumber(formatter, value, typeName);

        if (number instanceof BigDecimal) {
            return (BigDecimal) number;
        }

        requireFiniteForImport(number.doubleValue(), typeName);

        return new BigDecimal(number.toString());
    }

    /**
     * Avoids the widening noise produced when a float is widened to a double (e.g. 0.1f -> 0.10000000149011612).
     * Returns the double corresponding to the float's short decimal representation ({@link Float#toString}). (0.1f -> 0.1)
     *
     * @param value the float value to convert
     * @return the double matching the float's short decimal representation
     */
    public static double floatToPlainDouble(final float value) {

        return Double.parseDouble(Float.toString(value));
    }

    /**
     * Checks whether the value is within the range [minInclusive, maxInclusive] before narrowing it to an integer type.
     * If it is outside the range, throws a {@link PxlCellCodecException} to prevent a silent overflow (wraparound).
     * (If within range, the fractional part is truncated as before by the caller's narrowing conversion.)
     *
     * @param value        the value to check
     * @param minInclusive the inclusive lower bound
     * @param maxInclusive the inclusive upper bound
     * @param typeName     the target type name, used in the error message
     * @return the value unchanged when it is within range
     * @throws PxlCellCodecException when the value is outside the range
     */
    public static BigDecimal requireWithinRange(final BigDecimal value,
                                                final long minInclusive,
                                                final long maxInclusive,
                                                final String typeName)
            throws PxlCellCodecException {

        if (value.compareTo(BigDecimal.valueOf(minInclusive)) < 0
                || value.compareTo(BigDecimal.valueOf(maxInclusive)) > 0) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_NUMBER_OUT_OF_RANGE, String.valueOf(value.toPlainString()), typeName, String.valueOf(minInclusive), String.valueOf(maxInclusive)));
        }

        return value;
    }

    /**
     * Converts a {@link Number} to {@link BigDecimal} (via its {@code toString}, preserving the decimal magnitude) and then
     * delegates to {@link #requireWithinRange(BigDecimal, long, long, String)} to enforce the [minInclusive, maxInclusive] range.
     *
     * @param value        the value to check
     * @param minInclusive the inclusive lower bound
     * @param maxInclusive the inclusive upper bound
     * @param typeName     the target type name, used in the error message
     * @return the value as a {@link BigDecimal} when it is within range
     * @throws PxlCellCodecException when the value cannot be parsed as a number or is outside the range
     */
    public static BigDecimal requireWithinRange(final Number value,
                                                final long minInclusive,
                                                final long maxInclusive,
                                                final String typeName)
            throws PxlCellCodecException {

        final BigDecimal bigDecimalValue;
        try {
            bigDecimalValue = new BigDecimal(value.toString());
        } catch (NumberFormatException numberFormatException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_NUMBER_OUT_OF_RANGE, String.valueOf(value), typeName, String.valueOf(minInclusive), String.valueOf(maxInclusive)), numberFormatException);
        }

        return requireWithinRange(bigDecimalValue, minInclusive, maxInclusive, typeName);
    }

    /**
     * Validates a numeric-cell {@code double} against an integer range. Rejects non-finite values (NaN/Infinity) up front
     * - symmetric with the {@code float}/{@code double} import guard and preventing the bare {@link NumberFormatException}
     * that {@link BigDecimal#valueOf(double)} would otherwise throw - then enforces [minInclusive, maxInclusive] using exact
     * {@link BigDecimal} comparison (needed because bounds such as {@code Long.MAX_VALUE} are not exactly representable as a
     * {@code double}). Converting the {@code double} inside this method keeps {@link BigDecimal#valueOf(double)} off the call
     * sites, so a non-finite value can never leak an unchecked exception past this guard.
     *
     * @param value        the numeric cell value
     * @param minInclusive the inclusive lower bound
     * @param maxInclusive the inclusive upper bound
     * @param typeName     the target type name, used in the error message
     * @return the value as a {@link BigDecimal} when it is finite and within range
     * @throws PxlCellCodecException when the value is NaN/Infinity or outside the range
     */
    public static BigDecimal requireWithinRange(final double value,
                                                final long minInclusive,
                                                final long maxInclusive,
                                                final String typeName)
            throws PxlCellCodecException {

        requireFiniteForImport(value, typeName);
        return requireWithinRange(BigDecimal.valueOf(value), minInclusive, maxInclusive, typeName);
    }

    /**
     * Rejects non-finite values (NaN/Infinity) of a numeric-cell {@code double} and returns it as a {@link BigDecimal}.
     * For unbounded target types ({@link BigInteger}/{@link BigDecimal}) that have no fixed range: it applies the same
     * finiteness guard as {@link #requireWithinRange(double, long, long, String)} but skips the range check, again keeping
     * {@link BigDecimal#valueOf(double)} off the call sites.
     *
     * @param value    the numeric cell value
     * @param typeName the target type name, used in the error message
     * @return the value as a {@link BigDecimal} when it is finite
     * @throws PxlCellCodecException when the value is NaN or Infinity
     */
    public static BigDecimal requireFinite(final double value, final String typeName)
            throws PxlCellCodecException {

        requireFiniteForImport(value, typeName);
        return BigDecimal.valueOf(value);
    }

    /**
     * Rejects non-finite floating-point values (NaN, positive/negative Infinity) on import so that the import direction
     * stays symmetric with export, which also refuses them. A {@code null} value is allowed (a blank cell). A finite
     * {@code double} that overflows the {@code float} range narrows to Infinity and is therefore also rejected here.
     * <p>
     * A primitive {@code float} argument is autoboxed to this overload, so both {@link Float} and {@code float} codecs use it.
     *
     * @param value    the parsed value, or {@code null}
     * @param typeName the target type name, used in the error message
     * @throws PxlCellCodecException if the value is NaN or Infinity
     */
    public static void requireFiniteForImport(final Float value, final String typeName)
            throws PxlCellCodecException {

        if (Objects.nonNull(value) && (value.isNaN() || value.isInfinite())) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_NUMBER_NAN_READ, typeName, String.valueOf(value)));
        }
    }

    /**
     * Rejects non-finite floating-point values (NaN, positive/negative Infinity) on import so that the import direction
     * stays symmetric with export, which also refuses them. A {@code null} value is allowed (a blank cell).
     * <p>
     * A primitive {@code double} argument is autoboxed to this overload, so both {@link Double} and {@code double} codecs use it.
     *
     * @param value    the parsed value, or {@code null}
     * @param typeName the target type name, used in the error message
     * @throws PxlCellCodecException if the value is NaN or Infinity
     */
    public static void requireFiniteForImport(final Double value, final String typeName)
            throws PxlCellCodecException {

        if (Objects.nonNull(value) && (value.isNaN() || value.isInfinite())) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_NUMBER_NAN_READ, typeName, String.valueOf(value)));
        }
    }

    /**
     * Rejects non-finite floating-point values (NaN, positive/negative Infinity) on export, since they cannot be written as
     * a numeric cell (POI's {@code setCellValue(double)} would otherwise produce an error cell). A {@code null} value is allowed.
     * <p>
     * A primitive {@code float} argument is autoboxed to this overload, so both {@link Float} and {@code float} codecs use it.
     *
     * @param value the value about to be written, or {@code null}
     * @throws PxlCellCodecException if the value is NaN or Infinity
     */
    public static void requireFiniteForExport(final Float value)
            throws PxlCellCodecException {

        if (Objects.nonNull(value) && (value.isNaN() || value.isInfinite())) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_NUMBER_NAN_WRITE, String.valueOf(value)));
        }
    }

    /**
     * Rejects non-finite floating-point values (NaN, positive/negative Infinity) on export, since they cannot be written as
     * a numeric cell (POI's {@code setCellValue(double)} would otherwise produce an error cell). A {@code null} value is allowed.
     * <p>
     * A primitive {@code double} argument is autoboxed to this overload, so both {@link Double} and {@code double} codecs use it.
     *
     * @param value the value about to be written, or {@code null}
     * @throws PxlCellCodecException if the value is NaN or Infinity
     */
    public static void requireFiniteForExport(final Double value)
            throws PxlCellCodecException {

        if (Objects.nonNull(value) && (value.isNaN() || value.isInfinite())) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_NUMBER_NAN_WRITE, String.valueOf(value)));
        }
    }

}
