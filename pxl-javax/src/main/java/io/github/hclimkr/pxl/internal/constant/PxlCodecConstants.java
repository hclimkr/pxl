package io.github.hclimkr.pxl.internal.constant;

import io.github.hclimkr.pxl.internal.support.PxlDateTimeSupport;
import org.apache.poi.ss.usermodel.BuiltinFormats;

import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.stream.Collectors;

/**
 * Default date and time formatters and Excel display formats used by the date/time codecs.
 * <p>
 * The default patterns are fixed, locale-independent ISO-8601 patterns built with {@link Locale#ROOT}, so a
 * value handled without a column {@code pattern} is written and read the same way on every machine - nothing
 * here depends on the JVM default locale. Date-time write patterns use the ISO {@code 'T'} separator
 * ({@code yyyy-MM-dd'T'HH:mm:ss}, e.g. {@code 2026-07-23T08:45:58}); the read patterns are the padded shape
 * with lenient single-letter fields (accepting unpadded input) and accept both the {@code 'T'} separator and
 * a space, so values written before the ISO switch still round-trip. Cells written as Excel numeric date
 * serials use the fixed {@code *ExcelFormat} built-in format codes (also locale-independent).
 */
public final class PxlCodecConstants {

    // Write patterns: fixed ISO-8601, locale-independent. Date-time uses the ISO 'T' separator (e.g. 2026-07-23T08:45:58).
    private static final String javaDateWritePattern = "yyyy-MM-dd'T'HH:mm:ss";
    private static final String localDateWritePattern = "yyyy-MM-dd";
    private static final String localTimeWritePattern = "HH:mm:ss";
    private static final String localDateTimeWritePattern = "yyyy-MM-dd'T'HH:mm:ss";

    // Read patterns: lenient single-letter fields (accept unpadded input). Date-time accepts both the ISO 'T'
    // separator and a space, so values written before the ISO switch still round-trip.
    private static final String[] javaDateReadPatterns = {"y-M-d'T'H:m:s", "y-M-d'T'H:m", "y-M-d H:m:s", "y-M-d H:m"};
    private static final String[] localDateReadPatterns = {"y-M-d"};
    private static final String[] localTimeReadPatterns = {"H:m:s", "H:m"};
    private static final String[] localDateTimeReadPatterns = {"y-M-d'T'H:m:s", "y-M-d'T'H:m", "y-M-d H:m:s", "y-M-d H:m"};

    // SimpleDateFormat is not thread-safe, so a per-thread instance is provided.
    /**
     * Per-thread {@link Date} write formatter (fixed ISO pattern, {@link Locale#ROOT}).
     */
    public static final ThreadLocal<SimpleDateFormat> javaDateWriteFormatter =
            ThreadLocal.withInitial(() -> PxlDateTimeSupport.getCellSimpleDateFormatter(javaDateWritePattern, Locale.ROOT));

    /**
     * Per-thread {@link Date} read formatters (fixed ISO patterns, {@link Locale#ROOT}).
     */
    public static final ThreadLocal<List<SimpleDateFormat>> javaDateReadFormatters =
            ThreadLocal.withInitial(() -> Arrays.stream(javaDateReadPatterns)
                    .map(pattern -> PxlDateTimeSupport.getCellSimpleDateFormatter(pattern, Locale.ROOT))
                    .collect(Collectors.toList()));

    /**
     * {@link LocalDate} write formatter (fixed ISO pattern).
     */
    public static final DateTimeFormatter localDateWriteFormatter =
            PxlDateTimeSupport.getCellDateTimeFormatter(localDateWritePattern, Locale.ROOT);

    /**
     * {@link LocalDate} read formatters (fixed ISO patterns).
     */
    public static final List<DateTimeFormatter> localDateReadFormatters =
            Arrays.stream(localDateReadPatterns)
                    .map(pattern -> PxlDateTimeSupport.getCellDateTimeFormatter(pattern, Locale.ROOT))
                    .collect(Collectors.toList());

    /**
     * {@link LocalTime} write formatter (fixed ISO pattern).
     */
    public static final DateTimeFormatter localTimeWriteFormatter =
            PxlDateTimeSupport.getCellDateTimeFormatter(localTimeWritePattern, Locale.ROOT);

    /**
     * {@link LocalTime} read formatters (fixed ISO patterns).
     */
    public static final List<DateTimeFormatter> localTimeReadFormatters =
            Arrays.stream(localTimeReadPatterns)
                    .map(pattern -> PxlDateTimeSupport.getCellDateTimeFormatter(pattern, Locale.ROOT))
                    .collect(Collectors.toList());

    /**
     * {@link LocalDateTime} write formatter (fixed ISO pattern).
     */
    public static final DateTimeFormatter localDateTimeWriteFormatter =
            PxlDateTimeSupport.getCellDateTimeFormatter(localDateTimeWritePattern, Locale.ROOT);

    /**
     * {@link LocalDateTime} read formatters (fixed ISO patterns).
     */
    public static final List<DateTimeFormatter> localDateTimeReadFormatters =
            Arrays.stream(localDateTimeReadPatterns)
                    .map(pattern -> PxlDateTimeSupport.getCellDateTimeFormatter(pattern, Locale.ROOT))
                    .collect(Collectors.toList());

    // Excel display format codes applied when a date/time value is exported as a Numeric (Excel date serial) cell (POI built-in, locale-independent).
    /**
     * Excel display format code for {@link Date} numeric-serial cells.
     */
    public static final String javaDateExcelFormat = BuiltinFormats.getBuiltinFormat(14);

    /**
     * Excel display format code for {@link LocalDate} numeric-serial cells.
     */
    public static final String localDateExcelFormat = BuiltinFormats.getBuiltinFormat(14);

    /**
     * Excel display format code for {@link LocalTime} numeric-serial cells.
     */
    public static final String localTimeExcelFormat = BuiltinFormats.getBuiltinFormat(21);

    /**
     * Excel display format code for {@link LocalDateTime} numeric-serial cells.
     */
    public static final String localDateTimeExcelFormat = BuiltinFormats.getBuiltinFormat(22);

    /**
     * Prevents instantiation.
     */
    private PxlCodecConstants() {

        throw new AssertionError("no instances of this class");
    }

}
