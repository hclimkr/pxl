package io.github.hclimkr.pxl.internal.support;

import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.time.format.ResolverStyle;
import java.time.temporal.ChronoField;
import java.util.Date;
import java.util.Locale;
import java.util.Objects;

/**
 * Date/time formatting helpers shared by the date/time codecs.
 * <p>
 * Builds strict (non-lenient) formatters for pattern-based parsing/formatting of cell values, and converts between the
 * legacy {@link Date} and {@link LocalDateTime} using the system default time zone. Formatters are strict on purpose so
 * that out-of-range or malformed values fail rather than being silently rolled over.
 */
public final class PxlDateTimeSupport {

    private static final boolean dateLenient = false;
    private static final ResolverStyle resolverStyle = ResolverStyle.STRICT;

    /**
     * Prevents instantiation.
     */
    private PxlDateTimeSupport() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Creates a non-lenient {@link SimpleDateFormat} for the given pattern and locale (used for legacy {@link Date}
     * cells). The locale governs text fields such as month/day names and AM/PM markers; pass {@link Locale#ROOT} for a
     * locale-independent formatter and a specific locale to parse/format localized text.
     *
     * @param cellPattern the date/time pattern
     * @param locale      the locale for text fields
     * @return a non-lenient formatter for the pattern
     */
    public static SimpleDateFormat getCellSimpleDateFormatter(final String cellPattern,
                                                              final Locale locale) {

        final SimpleDateFormat simpleDateFormatter = new SimpleDateFormat(cellPattern, locale);
        simpleDateFormatter.setLenient(dateLenient);
        return simpleDateFormatter;
    }

    /**
     * Creates a strict {@link DateTimeFormatter} for the given pattern and locale, defaulting the {@code ERA} field to
     * CE (1) so that patterns without an era token still resolve under {@link ResolverStyle#STRICT}. The locale governs
     * text fields such as month/day names and AM/PM markers; pass {@link Locale#ROOT} for a locale-independent formatter
     * and a specific locale to parse/format localized text.
     *
     * @param cellPattern the date/time pattern
     * @param locale      the locale for text fields
     * @return a strict formatter for the pattern
     */
    public static DateTimeFormatter getCellDateTimeFormatter(final String cellPattern,
                                                             final Locale locale) {

        return new DateTimeFormatterBuilder().appendPattern(cellPattern)
                .parseDefaulting(ChronoField.ERA, 1)
                .toFormatter(locale)
                .withResolverStyle(resolverStyle);
    }

    /**
     * Converts a legacy {@link Date} to a {@link LocalDateTime} using the system default time zone.
     *
     * @param date the date to convert (may be {@code null})
     * @return the corresponding {@link LocalDateTime}, or {@code null} if {@code date} is {@code null}
     */
    public static LocalDateTime javaDateToLocalDateTime(final Date date) {

        if (Objects.isNull(date)) {
            return null;
        }

        return LocalDateTime.ofInstant(Instant.ofEpochMilli(date.getTime()), ZoneId.systemDefault());
    }

    /**
     * Converts a {@link LocalDateTime} to a legacy {@link Date} using the system default time zone.
     *
     * @param localDateTime the value to convert (may be {@code null})
     * @return the corresponding {@link Date}, or {@code null} if {@code localDateTime} is {@code null}
     */
    public static Date localDateTimeToJavaDate(final LocalDateTime localDateTime) {

        if (Objects.isNull(localDateTime)) {
            return null;
        }

        return Date.from(localDateTime.atZone(ZoneId.systemDefault()).toInstant());
    }

    /**
     * Converts a Java date/time pattern (as used by {@link SimpleDateFormat} / {@link DateTimeFormatter}) into the
     * equivalent Excel cell number-format code.
     * <p>
     * Excel writes every date/time field letter in lower case (month and minute are both {@code m}, disambiguated by
     * position relative to {@code h}/{@code s}), so the conversion simply lower-cases the pattern letters while leaving
     * separators untouched. Text inside single quotes is a Java literal section and is copied verbatim. Note that AM/PM
     * ({@code a}) markers are not translated, so they must not appear in patterns that are also exported as Excel numeric
     * date formats.
     *
     * @param javaPattern the Java date/time pattern
     * @return the corresponding Excel number-format code
     */
    public static String toExcelDateFormat(final String javaPattern) {

        final StringBuilder excelFormat = new StringBuilder(javaPattern.length());
        boolean inLiteral = false;

        for (int i = 0; i < javaPattern.length(); i++) {
            final char ch = javaPattern.charAt(i);
            if (ch == '\'') {
                inLiteral = !inLiteral;
                excelFormat.append(ch);
            } else if (inLiteral) {
                excelFormat.append(ch);
            } else {
                excelFormat.append(Character.toLowerCase(ch));
            }
        }

        return excelFormat.toString();
    }

}
