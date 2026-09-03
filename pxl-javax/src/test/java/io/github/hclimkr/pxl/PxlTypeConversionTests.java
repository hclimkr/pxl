package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.exception.PxlReflectionException;
import io.github.hclimkr.pxl.exception.PxlValidationException;
import io.github.hclimkr.pxl.tcdata.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInfo;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.time.*;
import java.util.*;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static io.github.hclimkr.pxl.tcdata.TestExports.emit;
import static io.github.hclimkr.pxl.tcdata.TestExports.workbookOf;
import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Per-type value conversion (codec) tests.
 * <p>
 * Round-trips each supported type through export->import with various values (boundary, negative, zero,
 * null, empty, special characters, collection position fidelity), and also verifies the import-only
 * conversion behavior of reading a type from an external cell (NUMERIC/BOOLEAN).
 */
public class PxlTypeConversionTests {

    private static Pxl pxl;

    @BeforeAll
    public static void setUpBeforeClass() {
        pxl = new Pxl();
    }

    // Captures the current test method name to match it with the export file name.
    private TestInfo testInfo;

    @BeforeEach
    public void bindTestInfo(final TestInfo testInfo) {
        this.testInfo = testInfo;
    }

    // Result of exporting an AllTypesRow list to a real file as a single sheet and importing it back. (file name = test method name)
    private List<AllTypesRow> roundTrip(final ExportDest dest, final List<AllTypesRow> rows) throws Exception {
        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(AllTypesRow.class, rows, "Types")
                .override(noValidationOption()), dest, testInfo);
        return pxl.importExcel()
                .sheet(AllTypesRow.class, Arrays.asList("Types"))
                .fromStream(new ByteArrayInputStream(bytes));
    }

    private AllTypesRow roundTrip1(final ExportDest dest, final AllTypesRow row) throws Exception {
        return roundTrip(dest, Arrays.asList(row)).get(0);
    }

    private static String repeat(final String unit, final int count) {
        return String.join("", Collections.nCopies(count, unit));
    }

    // A sheet built by composing specific cells directly (mimicking an external file).
    private interface SheetBuilder {
        void build(Sheet sheet);
    }

    private static byte[] sheet(final String sheetName, final SheetBuilder builder) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            builder.build(workbook.createSheet(sheetName));
            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    private static byte[] stringSheet(final String sheetName, final String[] headers, final String[][] dataRows) throws Exception {
        return sheet(sheetName, s -> {
            final Row header = s.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i).setCellValue(headers[i]);
            }
            for (int r = 0; r < dataRows.length; r++) {
                final Row row = s.createRow(r + 1);
                for (int c = 0; c < dataRows[r].length; c++) {
                    row.createCell(c).setCellValue(dataRows[r][c]);
                }
            }
        });
    }

    private static <T> List<T> importList(final byte[] bytes, final String sheetName, final Class<T> rowClass) throws Exception {
        return pxl.importExcel()
                .sheet(rowClass, Arrays.asList(sheetName))
                .fromStream(new ByteArrayInputStream(bytes));
    }

    // ==================================================================
    // Round-trip: various values per type
    // ==================================================================

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void integerTypes_boundaryValues_roundTrip(final ExportDest dest) throws Exception {
        final AllTypesRow max = Fixtures.baseAllTypesRow();
        max.setPrimByte(Byte.MAX_VALUE);
        max.setWrapByte(Byte.MAX_VALUE);
        max.setPrimShort(Short.MAX_VALUE);
        max.setWrapShort(Short.MAX_VALUE);
        max.setPrimInt(Integer.MAX_VALUE);
        max.setWrapInt(Integer.MAX_VALUE);
        max.setPrimLong(9007199254740991L);
        max.setWrapLong(9007199254740991L);

        final AllTypesRow min = Fixtures.baseAllTypesRow();
        min.setPrimByte(Byte.MIN_VALUE);
        min.setWrapByte(Byte.MIN_VALUE);
        min.setPrimShort(Short.MIN_VALUE);
        min.setWrapShort(Short.MIN_VALUE);
        min.setPrimInt(Integer.MIN_VALUE);
        min.setWrapInt(Integer.MIN_VALUE);
        min.setPrimLong(-9007199254740991L);
        min.setWrapLong(-9007199254740991L);

        final AllTypesRow zero = Fixtures.baseAllTypesRow();
        zero.setPrimByte((byte) 0);
        zero.setWrapByte((byte) 0);
        zero.setPrimInt(0);
        zero.setWrapLong(0L);

        final List<AllTypesRow> out = roundTrip(dest, Arrays.asList(max, min, zero));

        assertThat(out.get(0).getPrimByte()).isEqualTo(Byte.MAX_VALUE);
        assertThat(out.get(0).getWrapByte()).isEqualTo(Byte.MAX_VALUE);
        assertThat(out.get(0).getPrimShort()).isEqualTo(Short.MAX_VALUE);
        assertThat(out.get(0).getWrapShort()).isEqualTo(Short.MAX_VALUE);
        assertThat(out.get(0).getPrimInt()).isEqualTo(Integer.MAX_VALUE);
        assertThat(out.get(0).getWrapInt()).isEqualTo(Integer.MAX_VALUE);
        assertThat(out.get(0).getPrimLong()).isEqualTo(9007199254740991L);
        assertThat(out.get(0).getWrapLong()).isEqualTo(9007199254740991L);

        assertThat(out.get(1).getPrimByte()).isEqualTo(Byte.MIN_VALUE);
        assertThat(out.get(1).getWrapShort()).isEqualTo(Short.MIN_VALUE);
        assertThat(out.get(1).getPrimInt()).isEqualTo(Integer.MIN_VALUE);
        assertThat(out.get(1).getPrimLong()).isEqualTo(-9007199254740991L);
        assertThat(out.get(1).getWrapLong()).isEqualTo(-9007199254740991L);

        assertThat(out.get(2).getPrimByte()).isEqualTo((byte) 0);
        assertThat(out.get(2).getPrimInt()).isEqualTo(0);
        assertThat(out.get(2).getWrapLong()).isEqualTo(0L);
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void integerWrappers_null_roundTripsAsNull(final ExportDest dest) throws Exception {
        final AllTypesRow row = Fixtures.baseAllTypesRow();
        row.setWrapByte(null);
        row.setWrapShort(null);
        row.setWrapInt(null);
        row.setWrapLong(null);

        final AllTypesRow out = roundTrip1(dest, row);

        assertThat(out.getWrapByte()).isNull();
        assertThat(out.getWrapShort()).isNull();
        assertThat(out.getWrapInt()).isNull();
        assertThat(out.getWrapLong()).isNull();
        assertThat(out.getPrimInt()).isEqualTo(row.getPrimInt());
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void floatTypes_variousValues_roundTrip(final ExportDest dest) throws Exception {
        final AllTypesRow a = Fixtures.baseAllTypesRow();
        a.setPrimDouble(0.0);
        a.setWrapDouble(-2.25);
        a.setPrimFloat(0.0F);
        a.setWrapFloat(-4.5F);

        final AllTypesRow b = Fixtures.baseAllTypesRow();
        b.setPrimDouble(1234567.5);
        b.setWrapDouble(-9876.125);
        b.setPrimFloat(65504.0F);
        b.setWrapFloat(-1.5F);

        final AllTypesRow nulls = Fixtures.baseAllTypesRow();
        nulls.setWrapDouble(null);
        nulls.setWrapFloat(null);

        final List<AllTypesRow> out = roundTrip(dest, Arrays.asList(a, b, nulls));

        assertThat(out.get(0).getPrimDouble()).isEqualTo(0.0);
        assertThat(out.get(0).getWrapDouble()).isEqualTo(-2.25);
        assertThat(out.get(0).getPrimFloat()).isEqualTo(0.0F);
        assertThat(out.get(0).getWrapFloat()).isEqualTo(-4.5F);

        assertThat(out.get(1).getPrimDouble()).isEqualTo(1234567.5);
        assertThat(out.get(1).getWrapDouble()).isEqualTo(-9876.125);
        assertThat(out.get(1).getPrimFloat()).isEqualTo(65504.0F);
        assertThat(out.get(1).getWrapFloat()).isEqualTo(-1.5F);

        assertThat(out.get(2).getWrapDouble()).isNull();
        assertThat(out.get(2).getWrapFloat()).isNull();
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void bigNumbers_variousValues_roundTrip(final ExportDest dest) throws Exception {
        final AllTypesRow big = Fixtures.baseAllTypesRow();
        big.setBigInt(new BigInteger("123456789012345678901234567890"));
        big.setBigDec(new BigDecimal("3.14159265358979323846"));

        final AllTypesRow neg = Fixtures.baseAllTypesRow();
        neg.setBigInt(new BigInteger("-98765432109876543210"));
        neg.setBigDec(new BigDecimal("-0.0001"));

        final AllTypesRow zeroScale = Fixtures.baseAllTypesRow();
        zeroScale.setBigInt(BigInteger.ZERO);
        zeroScale.setBigDec(new BigDecimal("100.00"));

        final AllTypesRow nulls = Fixtures.baseAllTypesRow();
        nulls.setBigInt(null);
        nulls.setBigDec(null);

        final List<AllTypesRow> out = roundTrip(dest, Arrays.asList(big, neg, zeroScale, nulls));

        assertThat(out.get(0).getBigInt()).isEqualTo(new BigInteger("123456789012345678901234567890"));
        assertThat(out.get(0).getBigDec()).isEqualByComparingTo(new BigDecimal("3.14159265358979323846"));
        assertThat(out.get(1).getBigInt()).isEqualTo(new BigInteger("-98765432109876543210"));
        assertThat(out.get(1).getBigDec()).isEqualByComparingTo(new BigDecimal("-0.0001"));
        assertThat(out.get(2).getBigInt()).isEqualTo(BigInteger.ZERO);
        assertThat(out.get(2).getBigDec()).isEqualTo(new BigDecimal("100.00"));   // scale (2) preserved
        assertThat(out.get(3).getBigInt()).isNull();
        assertThat(out.get(3).getBigDec()).isNull();
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void boolean_trueFalseNull_roundTrip(final ExportDest dest) throws Exception {
        final AllTypesRow t = Fixtures.baseAllTypesRow();
        t.setPrimBool(true);
        t.setWrapBool(Boolean.TRUE);
        final AllTypesRow f = Fixtures.baseAllTypesRow();
        f.setPrimBool(false);
        f.setWrapBool(Boolean.FALSE);
        final AllTypesRow n = Fixtures.baseAllTypesRow();
        n.setWrapBool(null);

        final List<AllTypesRow> out = roundTrip(dest, Arrays.asList(t, f, n));

        assertThat(out.get(0).isPrimBool()).isTrue();
        assertThat(out.get(0).getWrapBool()).isTrue();
        assertThat(out.get(1).isPrimBool()).isFalse();
        assertThat(out.get(1).getWrapBool()).isFalse();
        assertThat(out.get(2).getWrapBool()).isNull();
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void char_variousValues_roundTrip(final ExportDest dest) throws Exception {
        final AllTypesRow letter = Fixtures.baseAllTypesRow();
        letter.setPrimChar('Z');
        letter.setWrapChar('a');
        final AllTypesRow digit = Fixtures.baseAllTypesRow();
        digit.setPrimChar('7');
        digit.setWrapChar('0');
        final AllTypesRow symbol = Fixtures.baseAllTypesRow();
        symbol.setPrimChar('#');
        symbol.setWrapChar('@');
        final AllTypesRow nullChar = Fixtures.baseAllTypesRow();
        nullChar.setWrapChar(null);

        final List<AllTypesRow> out = roundTrip(dest, Arrays.asList(letter, digit, symbol, nullChar));

        assertThat(out.get(0).getPrimChar()).isEqualTo('Z');
        assertThat(out.get(0).getWrapChar()).isEqualTo('a');
        assertThat(out.get(1).getPrimChar()).isEqualTo('7');
        assertThat(out.get(1).getWrapChar()).isEqualTo('0');
        assertThat(out.get(2).getPrimChar()).isEqualTo('#');
        assertThat(out.get(2).getWrapChar()).isEqualTo('@');
        assertThat(out.get(3).getWrapChar()).isNull();
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void string_specialCases_roundTrip(final ExportDest dest) throws Exception {
        final String longText = repeat("ab", 150);

        final AllTypesRow plain = Fixtures.baseAllTypesRow();
        plain.setText("Hello, World!");
        final AllTypesRow leadingZero = Fixtures.baseAllTypesRow();
        leadingZero.setText("007");
        final AllTypesRow innerSpace = Fixtures.baseAllTypesRow();
        innerSpace.setText("New York City");
        final AllTypesRow outerSpace = Fixtures.baseAllTypesRow();
        outerSpace.setText("  padded  ");
        final AllTypesRow onlySpaces = Fixtures.baseAllTypesRow();
        onlySpaces.setText("     ");
        final AllTypesRow empty = Fixtures.baseAllTypesRow();
        empty.setText("");
        final AllTypesRow symbols = Fixtures.baseAllTypesRow();
        symbols.setText("a@b#c$d%e^f&g*");
        final AllTypesRow longRow = Fixtures.baseAllTypesRow();
        longRow.setText(longText);

        final List<AllTypesRow> out = roundTrip(dest, Arrays.asList(
                plain, leadingZero, innerSpace, outerSpace, onlySpaces, empty, symbols, longRow));

        assertThat(out.get(0).getText()).isEqualTo("Hello, World!");
        assertThat(out.get(1).getText()).isEqualTo("007");
        assertThat(out.get(2).getText()).isEqualTo("New York City");
        assertThat(out.get(3).getText()).isEqualTo("padded");   // import trim
        assertThat(out.get(4).getText()).isNull();              // whitespace only -> null
        assertThat(out.get(5).getText()).isNull();              // empty -> null
        assertThat(out.get(6).getText()).isEqualTo("a@b#c$d%e^f&g*");
        assertThat(out.get(7).getText()).isEqualTo(longText);
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void enum_allConstantsAndNull_roundTrip(final ExportDest dest) throws Exception {
        final List<AllTypesRow> rows = new ArrayList<>();
        for (final Grade grade : Grade.values()) {
            final AllTypesRow row = Fixtures.baseAllTypesRow();
            row.setGrade(grade);
            rows.add(row);
        }
        for (final Category category : Category.values()) {
            final AllTypesRow row = Fixtures.baseAllTypesRow();
            row.setCategory(category);
            rows.add(row);
        }
        final AllTypesRow nullEnums = Fixtures.baseAllTypesRow();
        nullEnums.setGrade(null);
        nullEnums.setCategory(null);
        rows.add(nullEnums);

        final List<AllTypesRow> out = roundTrip(dest, rows);

        final Grade[] grades = Grade.values();
        for (int i = 0; i < grades.length; i++) {
            assertThat(out.get(i).getGrade()).isEqualTo(grades[i]);
        }
        final Category[] categories = Category.values();
        for (int i = 0; i < categories.length; i++) {
            assertThat(out.get(grades.length + i).getCategory()).isEqualTo(categories[i]);
        }
        final AllTypesRow last = out.get(out.size() - 1);
        assertThat(last.getGrade()).isNull();
        assertThat(last.getCategory()).isNull();
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void dateTime_variousValues_roundTrip(final ExportDest dest) throws Exception {
        final ZoneId zone = ZoneId.systemDefault();

        final LocalDateTime midnight = LocalDateTime.of(2000, 1, 1, 0, 0, 0);
        final AllTypesRow a = Fixtures.baseAllTypesRow();
        a.setJavaDate(Date.from(midnight.atZone(zone).toInstant()));
        a.setLocalDate(LocalDate.of(2000, 1, 1));
        a.setLocalTime(LocalTime.of(0, 0, 0));
        a.setLocalDateTime(midnight);
        a.setZonedDateTime(midnight.atZone(zone));
        a.setOffsetDateTime(midnight.atZone(zone).toOffsetDateTime());
        a.setDuration(Duration.ZERO);
        a.setPeriod(Period.ZERO);

        final LocalDateTime endOfDay = LocalDateTime.of(2038, 12, 31, 23, 59, 59);
        final AllTypesRow b = Fixtures.baseAllTypesRow();
        b.setJavaDate(Date.from(endOfDay.atZone(zone).toInstant()));
        b.setLocalDate(LocalDate.of(2038, 12, 31));
        b.setLocalTime(LocalTime.of(23, 59, 59));
        b.setLocalDateTime(endOfDay);
        b.setZonedDateTime(endOfDay.atZone(zone));
        b.setOffsetDateTime(endOfDay.atZone(zone).toOffsetDateTime());
        b.setDuration(Duration.ofDays(2).plusHours(3).plusMinutes(4).plusSeconds(5));
        b.setPeriod(Period.of(2, 6, 15));

        final AllTypesRow nulls = Fixtures.baseAllTypesRow();
        nulls.setJavaDate(null);
        nulls.setLocalDate(null);
        nulls.setLocalTime(null);
        nulls.setLocalDateTime(null);
        nulls.setZonedDateTime(null);
        nulls.setOffsetTime(null);
        nulls.setOffsetDateTime(null);
        nulls.setDuration(null);
        nulls.setPeriod(null);

        final List<AllTypesRow> out = roundTrip(dest, Arrays.asList(a, b, nulls));

        assertThat(out.get(0).getJavaDate()).isEqualTo(a.getJavaDate());
        assertThat(out.get(0).getLocalDate()).isEqualTo(LocalDate.of(2000, 1, 1));
        assertThat(out.get(0).getLocalTime()).isEqualTo(LocalTime.of(0, 0, 0));
        assertThat(out.get(0).getLocalDateTime()).isEqualTo(midnight);
        assertThat(out.get(0).getZonedDateTime()).isEqualTo(midnight.atZone(zone));
        assertThat(out.get(0).getOffsetDateTime()).isEqualTo(midnight.atZone(zone).toOffsetDateTime());
        assertThat(out.get(0).getDuration()).isEqualTo(Duration.ZERO);
        assertThat(out.get(0).getPeriod()).isEqualTo(Period.ZERO);

        assertThat(out.get(1).getJavaDate()).isEqualTo(b.getJavaDate());
        assertThat(out.get(1).getLocalDate()).isEqualTo(LocalDate.of(2038, 12, 31));
        assertThat(out.get(1).getLocalTime()).isEqualTo(LocalTime.of(23, 59, 59));
        assertThat(out.get(1).getLocalDateTime()).isEqualTo(endOfDay);
        assertThat(out.get(1).getDuration()).isEqualTo(Duration.ofDays(2).plusHours(3).plusMinutes(4).plusSeconds(5));
        assertThat(out.get(1).getPeriod()).isEqualTo(Period.of(2, 6, 15));

        assertThat(out.get(2).getJavaDate()).isNull();
        assertThat(out.get(2).getLocalDateTime()).isNull();
        assertThat(out.get(2).getZonedDateTime()).isNull();
        assertThat(out.get(2).getOffsetDateTime()).isNull();
        assertThat(out.get(2).getDuration()).isNull();
        assertThat(out.get(2).getPeriod()).isNull();
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void customObject_variousValues_roundTrip(final ExportDest dest) throws Exception {
        final AllTypesRow a = Fixtures.baseAllTypesRow();
        a.setPoint(new Point(-3, -7));
        a.setMoney(new Money("EUR", 999999999L));
        final AllTypesRow b = Fixtures.baseAllTypesRow();
        b.setPoint(new Point(0, 0));
        b.setMoney(new Money("JPY", 0L));
        final AllTypesRow nulls = Fixtures.baseAllTypesRow();
        nulls.setPoint(null);
        nulls.setMoney(null);

        final List<AllTypesRow> out = roundTrip(dest, Arrays.asList(a, b, nulls));

        assertThat(out.get(0).getPoint().getX()).isEqualTo(-3);
        assertThat(out.get(0).getPoint().getY()).isEqualTo(-7);
        assertThat(out.get(0).getMoney().getCurrency()).isEqualTo("EUR");
        assertThat(out.get(0).getMoney().getAmount()).isEqualTo(999999999L);
        assertThat(out.get(1).getPoint().getX()).isEqualTo(0);
        assertThat(out.get(1).getMoney().getAmount()).isEqualTo(0L);
        assertThat(out.get(2).getPoint()).isNull();
        assertThat(out.get(2).getMoney()).isNull();
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void collection_positionFidelity_roundTrip(final ExportDest dest) throws Exception {
        final AllTypesRow multi = Fixtures.baseAllTypesRow();
        multi.setStringList(Arrays.asList("Apple", "Banana", "Cherry"));
        multi.setIntList(Arrays.asList(10, 20, 30));
        multi.setGradeList(Arrays.asList(Grade.A, Grade.B, Grade.F));
        final AllTypesRow single = Fixtures.baseAllTypesRow();
        single.setStringList(Arrays.asList("Solo"));
        single.setIntList(Arrays.asList(42));
        single.setGradeList(Arrays.asList(Grade.C));
        final AllTypesRow innerNull = Fixtures.baseAllTypesRow();
        innerNull.setStringList(Arrays.asList("A", null, "C"));
        innerNull.setIntList(Arrays.asList(1, null, 3));
        innerNull.setGradeList(Arrays.asList(Grade.A, null, Grade.F));
        final AllTypesRow edgeNull = Fixtures.baseAllTypesRow();
        edgeNull.setStringList(Arrays.asList(null, "B", null));

        final List<AllTypesRow> out = roundTrip(dest, Arrays.asList(multi, single, innerNull, edgeNull));

        assertThat(out.get(0).getStringList()).containsExactly("Apple", "Banana", "Cherry");
        assertThat(out.get(0).getIntList()).containsExactly(10, 20, 30);
        assertThat(out.get(0).getGradeList()).containsExactly(Grade.A, Grade.B, Grade.F);
        assertThat(out.get(1).getStringList()).containsExactly("Solo");
        assertThat(out.get(1).getGradeList()).containsExactly(Grade.C);
        assertThat(out.get(2).getStringList()).containsExactly("A", null, "C");
        assertThat(out.get(2).getIntList()).containsExactly(1, null, 3);
        assertThat(out.get(2).getGradeList()).containsExactly(Grade.A, null, Grade.F);
        assertThat(out.get(3).getStringList()).containsExactly(null, "B", null);
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void collection_emptyAndNull_roundTripsAsNull(final ExportDest dest) throws Exception {
        final AllTypesRow empty = Fixtures.baseAllTypesRow();
        empty.setStringList(new ArrayList<>());
        empty.setIntList(new ArrayList<>());
        empty.setGradeList(new ArrayList<>());
        final AllTypesRow nulls = Fixtures.baseAllTypesRow();
        nulls.setStringList(null);
        nulls.setIntList(null);
        nulls.setGradeList(null);

        final List<AllTypesRow> out = roundTrip(dest, Arrays.asList(empty, nulls));

        // An empty collection is exported as an empty cell and becomes null on import.
        assertThat(out.get(0).getStringList()).isNull();
        assertThat(out.get(0).getIntList()).isNull();
        assertThat(out.get(0).getGradeList()).isNull();
        assertThat(out.get(1).getStringList()).isNull();
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void allNullableFields_null_roundTrip(final ExportDest dest) throws Exception {
        final AllTypesRow row = Fixtures.baseAllTypesRow();
        row.setText(null);
        row.setLeadingZero(null);
        row.setWrapByte(null);
        row.setWrapShort(null);
        row.setWrapInt(null);
        row.setWrapLong(null);
        row.setWrapDouble(null);
        row.setWrapFloat(null);
        row.setWrapChar(null);
        row.setWrapBool(null);
        row.setBigInt(null);
        row.setBigDec(null);
        row.setJavaDate(null);
        row.setLocalDate(null);
        row.setLocalTime(null);
        row.setLocalDateTime(null);
        row.setZonedDateTime(null);
        row.setOffsetTime(null);
        row.setOffsetDateTime(null);
        row.setDuration(null);
        row.setPeriod(null);
        row.setGrade(null);
        row.setCategory(null);
        row.setPoint(null);
        row.setMoney(null);
        row.setStringList(null);
        row.setIntList(null);
        row.setGradeList(null);

        final AllTypesRow out = roundTrip1(dest, row);

        assertThat(out.getText()).isNull();
        assertThat(out.getWrapInt()).isNull();
        assertThat(out.getBigDec()).isNull();
        assertThat(out.getJavaDate()).isNull();
        assertThat(out.getLocalDateTime()).isNull();
        assertThat(out.getGrade()).isNull();
        assertThat(out.getPoint()).isNull();
        assertThat(out.getMoney()).isNull();
        assertThat(out.getStringList()).isNull();
        // Primitive fields keep their base values
        assertThat(out.getPrimInt()).isEqualTo(5);
        assertThat(out.isPrimBool()).isTrue();
        assertThat(out.getPrimChar()).isEqualTo('A');
    }

    // ==================================================================
    // External cell (NUMERIC/BOOLEAN) -> type import conversion
    // ==================================================================

    @Test
    public void char_numericAndBooleanCell_takesFirstChar() throws Exception {
        // NUMERIC 12 -> first character '1' of "12"
        assertThat(importList(sheet("D", s -> {
            s.createRow(0).createCell(0).setCellValue("C");
            s.createRow(1).createCell(0).setCellValue(12);
        }), "D", CharRow.class).get(0).getC()).isEqualTo('1');
        // NUMERIC -3 -> '-'  (first character of "-3")
        assertThat(importList(sheet("D", s -> {
            s.createRow(0).createCell(0).setCellValue("C");
            s.createRow(1).createCell(0).setCellValue(-3);
        }), "D", CharRow.class).get(0).getC()).isEqualTo('-');
        // BOOLEAN true/false -> '1'/'0'  (first character of "1"/"0")
        assertThat(importList(sheet("D", s -> {
            s.createRow(0).createCell(0).setCellValue("C");
            s.createRow(1).createCell(0).setCellValue(true);
        }), "D", CharRow.class).get(0).getC()).isEqualTo('1');
        assertThat(importList(sheet("D", s -> {
            s.createRow(0).createCell(0).setCellValue("C");
            s.createRow(1).createCell(0).setCellValue(false);
        }), "D", CharRow.class).get(0).getC()).isEqualTo('0');
    }

    @Test
    public void primitiveChar_emptyValue_isJavaDefault() throws Exception {
        // An empty value leaves a char field at the Java default, the way the other primitives already behaved. A CSV
        // empty token used to bind a space (0x20), so a blank Excel cell (which the resolver drops before the codec,
        // leaving the field untouched) and an empty CSV field disagreed on the same absent value (issue L5 fix).
        final byte[] excelBytes = stringSheet("E", new String[]{"PrimChar"}, new String[][]{{""}});
        assertThat(importList(excelBytes, "E", AllTypesRow.class).get(0).getPrimChar()).isEqualTo((char) 0);

        final List<AllTypesRow> csvRows = pxl.importCsv()
                .sheet(AllTypesRow.class)
                .fromStream("E", new ByteArrayInputStream("PrimChar\n\"\"\n".getBytes(StandardCharsets.UTF_8)));
        assertThat(csvRows.get(0).getPrimChar()).isEqualTo((char) 0);
    }

    @Test
    public void boolean_numericCell_nonZeroIsTrue() throws Exception {
        // A numeric cell is true when it is not zero, which is the rule the REFERENCE states. The comparison used
        // to be against an epsilon (|x| > 1e-7), so a genuine small value such as 1e-8 read as false with nothing
        // to show for it (issue L3 fix). Both signed zeros stay false.
        final double[] numericValues = {1e-8, 2, -1, 0, -0.0};
        final byte[] bytes = sheet("B", s -> {
            s.createRow(0).createCell(0).setCellValue("Bool");
            for (int rowIndex = 0; rowIndex < numericValues.length; rowIndex++) {
                s.createRow(rowIndex + 1).createCell(0).setCellValue(numericValues[rowIndex]);
            }
        });

        assertThat(importList(bytes, "B", TypedRow.class)).extracting(TypedRow::getBool)
                .containsExactly(Boolean.TRUE, Boolean.TRUE, Boolean.TRUE, Boolean.FALSE, Boolean.FALSE);
    }

    @Test
    public void boolean_builtInTokens_parse() throws Exception {
        final byte[] bytes = stringSheet("B", new String[]{"Bool"},
                new String[][]{{"t"}, {"off"}, {"1"}, {"No"}, {"TRUE"}});
        assertThat(importList(bytes, "B", TypedRow.class)).extracting(TypedRow::getBool)
                .containsExactly(Boolean.TRUE, Boolean.FALSE, Boolean.TRUE, Boolean.FALSE, Boolean.TRUE);
    }

    @Test
    public void string_booleanCell_rendersViaImportTrueString() throws Exception {
        final byte[] bytes = sheet("F", s -> {
            s.createRow(0).createCell(0).setCellValue("Flag");
            s.createRow(1).createCell(0).setCellValue(true);
            s.createRow(2).createCell(0).setCellValue(false);
        });
        final List<StringFlagRow> rows = importList(bytes, "F", StringFlagRow.class);
        assertThat(rows.get(0).getFlag()).isEqualTo("YES");
        assertThat(rows.get(1).getFlag()).isEqualTo("NO");
    }

    @Test
    public void string_numericCellGroupingFormat_localeIndependent() throws Exception {
        // A NUMERIC cell read into a String is rendered with a Locale.ROOT DataFormatter, so decimal/grouping
        // symbols do not depend on the JVM default locale. Under de_DE (',' decimal / '.' grouping) the result
        // is still the ROOT form "2,012,000,046.50", not "2.012.000.046,50".
        final Locale previous = Locale.getDefault();
        try {
            Locale.setDefault(Locale.GERMANY);

            final byte[] bytes = sheet("N", s -> {
                s.createRow(0).createCell(0).setCellValue("Text");   // AllTypesRow "Text" is a String column
                final Workbook wb = s.getWorkbook();
                final CellStyle grouping = wb.createCellStyle();
                grouping.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
                final Cell cell = s.createRow(1).createCell(0);
                cell.setCellValue(2012000046.5);
                cell.setCellStyle(grouping);
            });

            final AllTypesRow row = importList(bytes, "N", AllTypesRow.class).get(0);
            assertThat(row.getText()).isEqualTo("2,012,000,046.50");
        } finally {
            Locale.setDefault(previous);
        }
    }

    @Test
    public void zonedDateTime_isoOffset_isPreserved() throws Exception {
        final byte[] bytes = stringSheet("Z", new String[]{"Zoned", "Offset"},
                new String[][]{{"2024-01-01T10:00:00+05:00", "2024-01-01T10:00:00+05:00"}});
        final ZonedRow row = importList(bytes, "Z", ZonedRow.class).get(0);
        assertThat(row.getZoned()).isEqualTo(ZonedDateTime.parse("2024-01-01T10:00:00+05:00"));
        assertThat(row.getOffset()).isEqualTo(OffsetDateTime.parse("2024-01-01T10:00:00+05:00"));
    }

    @Test
    public void collection_nestedGeneric_throws() throws Exception {
        final byte[] bytes = stringSheet("N", new String[]{"Nested"}, new String[][]{{"a;b"}});
        assertThrows(PxlReflectionException.class, () -> importList(bytes, "N", NestedCollectionRow.class));
    }

    @Test
    public void date_importPatternMismatch_fallsBackToDefault() throws Exception {
        // importPattern="dd/MM/yyyy" but ISO "2020-02-01" -> falls back and passes (import asymmetry)
        final byte[] bytes = stringSheet("L", new String[]{"D"}, new String[][]{{"2020-02-01"}});
        assertThat(importList(bytes, "L", LenientDateRow.class).get(0).getD()).isEqualTo(LocalDate.of(2020, 2, 1));
    }

    @Test
    public void duration_numericCell_truncatesFraction() throws Exception {
        final byte[] bytes = sheet("T", s -> {
            s.createRow(0).createCell(0).setCellValue("Dur");
            s.createRow(1).createCell(0).setCellValue(90.9);   // seconds, fraction truncated
        });
        assertThat(importList(bytes, "T", TypedRow.class).get(0).getDur()).isEqualTo(Duration.ofSeconds(90));
    }

    @Test
    public void period_numericOutOfRange_throws() throws Exception {
        final byte[] bytes = sheet("T", s -> {
            s.createRow(0).createCell(0).setCellValue("Per");
            s.createRow(1).createCell(0).setCellValue(3_000_000_000.0);   // exceeds int (days) range
        });
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "T", TypedRow.class));
    }

    @Test
    public void durationPeriod_customPattern_parsesAcrossRows() throws Exception {
        // Duration/Period columns with a custom DurationFormatUtils-style importPattern; the compiled pattern is reused across every row.
        final byte[] bytes = stringSheet("TP",
                new String[]{"Dur", "Per"},
                new String[][]{
                        {"01:02:03", "01/02/03"},
                        {"10:20:30", "05/06/07"},
                });

        final List<TemporalPatternRow> out = importList(bytes, "TP", TemporalPatternRow.class);

        assertThat(out).hasSize(2);
        assertThat(out.get(0).getDur()).isEqualTo(Duration.ofHours(1).plusMinutes(2).plusSeconds(3));
        assertThat(out.get(0).getPer()).isEqualTo(Period.of(1, 2, 3));
        assertThat(out.get(1).getDur()).isEqualTo(Duration.ofHours(10).plusMinutes(20).plusSeconds(30));
        assertThat(out.get(1).getPer()).isEqualTo(Period.of(5, 6, 7));
    }

    @Test
    public void durationPeriod_customPattern_isoFallbackOnMismatch() throws Exception {
        // Values are ISO-8601, not matching the DurationFormatUtils-style importPattern -> the parse falls back to ISO-8601.
        final byte[] bytes = stringSheet("TP",
                new String[]{"Dur", "Per"},
                new String[][]{{"PT2H30M", "P1Y6M"}});

        final TemporalPatternRow row = importList(bytes, "TP", TemporalPatternRow.class).get(0);

        assertThat(row.getDur()).isEqualTo(Duration.ofHours(2).plusMinutes(30));
        assertThat(row.getPer()).isEqualTo(Period.of(1, 6, 0));
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void long_withPattern_preservesLargeValue(final ExportDest dest) throws Exception {
        final long big = 9007199254740993L;   // 2^53 + 1

        final LongPatternRow row = new LongPatternRow();
        row.setBig(big);

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(LongPatternRow.class, Arrays.asList(row), "Big")
                .override(noValidationOption()), dest, testInfo);

        final LongPatternRow out = pxl.importExcel()
                .sheet(LongPatternRow.class, Arrays.asList("Big"))
                .fromStream(new ByteArrayInputStream(bytes)).get(0);
        assertThat(out.getBig()).isEqualTo(big);
    }

    // ==================================================================
    // Constraint / asymmetry behavior characterization (regression guard pinning current behavior)
    // ==================================================================

    @Test
    public void byteSize_withinRange_imports() throws Exception {
        // "abc" = 3 bytes <= max(5) -> passes (public constraint @PxlByteSize)
        final byte[] bytes = stringSheet("BS", new String[]{"Code"}, new String[][]{{"abc"}});
        assertThat(importList(bytes, "BS", ByteSizeRow.class).get(0).getCode()).isEqualTo("abc");
    }

    @Test
    public void byteSize_underMin_throws() throws Exception {
        // "abc" = 3 bytes < min(4) -> @PxlByteSize lower-bound violation
        final byte[] bytes = stringSheet("BS", new String[]{"Name"}, new String[][]{{"abc"}});
        assertThrows(PxlValidationException.class, () -> importList(bytes, "BS", ByteSizeRow.class));
    }

    @Test
    public void byteSize_charset_countsBytesInGivenCharset() throws Exception {
        // "안녕" = 4 bytes in EUC-KR (<=4) -> passes. (would be 6 bytes and exceed in UTF-8) - verifies byte counting per the given charset
        final byte[] bytes = stringSheet("BS", new String[]{"Label"}, new String[][]{{"안녕"}});
        assertThat(importList(bytes, "BS", ByteSizeRow.class).get(0).getLabel()).isEqualTo("안녕");
    }

    @Test
    public void byteSize_null_isValid() throws Exception {
        // Leave the constrained Name cell empty (-> null) and put a value in Code so the whole row is not ignored.
        // empty cell -> null -> still valid for @PxlByteSize(min=4) (null allowed)
        final byte[] bytes = stringSheet("BS", new String[]{"Code", "Name"}, new String[][]{{"ab", ""}});
        final ByteSizeRow row = importList(bytes, "BS", ByteSizeRow.class).get(0);
        assertThat(row.getName()).isNull();
        assertThat(row.getCode()).isEqualTo("ab");
    }

    @Test
    public void float_importOverflow_throws() throws Exception {
        // Reading an oversized value (1E300) from an external NUMERIC cell as float makes the (float) narrowing Infinity.
        // To mirror export's NaN/Infinity fail-fast rejection, import also rejects it with an exception (issue M4 fix).
        final byte[] bytes = sheet("O", s -> {
            s.createRow(0).createCell(0).setCellValue("WrapFloat");
            s.createRow(1).createCell(0).setCellValue(1.0E300);
        });
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "O", AllTypesRow.class));
    }

    @Test
    public void float_importInfinityString_throws() throws Exception {
        // The string "Infinity" is accepted by Float.parseFloat, but import rejects it too, symmetric with export (issue M4 fix).
        final byte[] bytes = stringSheet("O", new String[]{"WrapFloat"}, new String[][]{{"Infinity"}});
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "O", AllTypesRow.class));
    }

    @Test
    public void double_importNaNString_throws() throws Exception {
        // The string "NaN" is accepted by Double.parseDouble, but import rejects it too, symmetric with export (issue M4 fix).
        final byte[] bytes = stringSheet("N", new String[]{"WrapDouble"}, new String[][]{{"NaN"}});
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "N", AllTypesRow.class));
    }

    // ------------------------------------------------------------------
    // Non-finite (Infinity) NUMERIC cells: the 8 integer types, Duration, and Period are rejected by the finiteness guard (issue L2 fix)
    // A NUMERIC cell holding a literal beyond the double range makes getNumericCellValue() return Infinity;
    // before the guard, BigDecimal.valueOf(Infinity) threw a message-less NFE that lost the cause explanation.
    // This mirrors how the floating types (float/double) reject via requireFiniteForImport.
    // ------------------------------------------------------------------

    // Builds a NUMERIC cell whose getNumericCellValue() becomes Infinity by putting a literal beyond the double range ("1E309") as the raw value.
    // (setCellValue(Double.POSITIVE_INFINITY) is turned into an error cell by POI and never reaches the NUMERIC path, so the low-level CTCell is used.)
    private static byte[] numericInfinitySheet(final String header) throws Exception {
        return sheet("Inf", s -> {
            s.createRow(0).createCell(0).setCellValue(header);
            final XSSFCell cell =
                    (XSSFCell) s.createRow(1).createCell(0);
            cell.setCellValue(0.0);            // first make it a NUMERIC cell, then
            cell.getCTCell().setV("1E309");    // replace the raw value with a literal beyond the double range
        });
    }

    @Test
    public void numericTypes_nonFiniteNumericCell_throws() throws Exception {
        // Verify that a non-finite NUMERIC cell is rejected with PxlCellCodecException in each column of
        // the 8 integer types (primitive/wrapper), Duration, Period, BigInteger, and BigDecimal.
        final String[] headers = {
                "PrimByte", "WrapByte", "PrimShort", "WrapShort",
                "PrimInt", "WrapInt", "PrimLong", "WrapLong",
                "Duration", "Period", "BigInt", "BigDec"};
        for (final String header : headers) {
            final byte[] bytes = numericInfinitySheet(header);
            assertThrows(PxlCellCodecException.class,
                    () -> importList(bytes, "Inf", AllTypesRow.class),
                    header + " column: a non-finite NUMERIC cell must be rejected with PxlCellCodecException");
        }
    }

    @Test
    public void dateTimeTypes_negativeSerialNumericCell_throws() throws Exception {
        // POI answers null for a serial number outside the Excel date range (isValidExcelDate(-1) is false). The
        // java.time codecs used to dereference that null straight away, so the failure reached the caller as a
        // message-less NPE wrapped in PxlCellCodecException, while the Date codec assigned the null and bound
        // nothing at all. Every date/time column must reject the cell with a message that names the value.
        final String[] headers = {
                "JavaDate", "LocalDate", "LocalTime", "LocalDateTime",
                "ZonedDateTime", "OffsetTime", "OffsetDateTime"};
        for (final String header : headers) {
            final byte[] bytes = sheet("Neg", s -> {
                s.createRow(0).createCell(0).setCellValue(header);
                s.createRow(1).createCell(0).setCellValue(-1.0);   // plain numeric cell -> read as a raw Excel serial
            });

            final PxlCellCodecException exception = assertThrows(PxlCellCodecException.class,
                    () -> importList(bytes, "Neg", AllTypesRow.class),
                    header + " column: a negative Excel serial must be rejected with PxlCellCodecException");
            assertThat(exception)
                    .as(header + " column: the message must name the offending value instead of hiding behind an NPE")
                    .hasMessageContaining("-1");
        }
    }

    @Test
    public void integer_numericOutOfRange_throws() throws Exception {
        // A finite NUMERIC value beyond the Integer range -> the unified requireWithinRange(double, ...) passes finiteness then rejects via range check.
        // (The Period path is covered by period_numericOutOfRange_throws - here we verify the range check is retained for the integer codec family.)
        final byte[] bytes = sheet("O", s -> {
            s.createRow(0).createCell(0).setCellValue("WrapInt");
            s.createRow(1).createCell(0).setCellValue(3_000_000_000.0);   // > Integer.MAX_VALUE, finite
        });
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "O", AllTypesRow.class));
    }

    @Test
    public void bigTypes_largeFiniteNumericCell_beyondLongRange_imports() throws Exception {
        // requireFiniteAsBigDecimal has no range check - a finite NUMERIC value far beyond the Long range (~9.2e18) must also
        // import correctly as BigInteger/BigDecimal (if the unification misroutes Big* to the range overload, this test breaks).
        final byte[] bytes = sheet("Big", s -> {
            final Row header = s.createRow(0);
            header.createCell(0).setCellValue("BigInt");
            header.createCell(1).setCellValue("BigDec");
            final Row data = s.createRow(1);
            data.createCell(0).setCellValue(1.0E30);   // finite, far beyond Long.MAX
            data.createCell(1).setCellValue(1.0E30);
        });

        final AllTypesRow row = importList(bytes, "Big", AllTypesRow.class).get(0);
        assertThat(row.getBigInt()).isEqualTo(BigInteger.TEN.pow(30));                  // 10^30
        assertThat(row.getBigDec()).isEqualByComparingTo(new BigDecimal("1E30"));
    }

    @Test
    public void date_invalidDay_customPattern_throws() throws Exception {
        // java.util.Date + a custom importPattern becomes non-lenient and rejects an invalid date (2023-02-30) (issue M1 fix).
        // Previously it was lenient and silently rolled over to 2023-03-02. The same string also throws for LocalDate since it is strict.
        final byte[] bytes = stringSheet("R", new String[]{"D"}, new String[][]{{"2023-02-30"}});
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "R", LenientJavaDateRow.class));
    }

    @Test
    public void date_invalidMonth_customPattern_throws() throws Exception {
        // An out-of-range month (2023-13-01) is also rejected without rollover (-> 2024-01-01) (issue M1 fix).
        final byte[] bytes = stringSheet("R", new String[]{"D"}, new String[][]{{"2023-13-01"}});
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "R", LenientJavaDateRow.class));
    }

    @Test
    public void numeric_customPattern_trailingGarbage_throws() throws Exception {
        // A numeric custom pattern must consume the whole string, so "123abc" is rejected rather than read as 123
        // (issue M3 fix). This used to bind 123 silently, which made a patterned column more permissive than an
        // unpatterned one - the no-pattern path has always rejected the same value via Integer.parseInt and friends.
        final byte[] bytes = stringSheet("G", new String[]{"Amount"}, new String[][]{{"123abc"}});
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "G", PatternRow.class));
    }

    @Test
    public void numeric_customPattern_scientificNotation_throws() throws Exception {
        // "1e3" is the dangerous shape of the same defect: "#,##0" reads the leading 1 and stops, so the value used to
        // bind as 1 - a plausible number that no later validation could flag. The whole string must match (issue M3 fix).
        final byte[] bytes = stringSheet("G", new String[]{"WrapInt"}, new String[][]{{"1e3"}});
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "G", NumberPatternRow.class));
    }

    @Test
    public void numeric_customPattern_groupedValue_parses() throws Exception {
        // Requiring full consumption must not reject what the pattern legitimately reads: the grouping separator and
        // the decimal part are part of the pattern, so both values still bind.
        final byte[] bytes = stringSheet("G", new String[]{"WrapInt", "BigDec"}, new String[][]{{"1,234", "9,876.25"}});

        final NumberPatternRow row = importList(bytes, "G", NumberPatternRow.class).get(0);
        assertThat(row.getWrapInt()).isEqualTo(1234);
        assertThat(row.getBigDec()).isEqualByComparingTo(new BigDecimal("9876.25"));
    }

    @Test
    public void numeric_customPattern_infinityToken_throws() throws Exception {
        // DecimalFormat consumes its own infinity symbol in full, so full-consumption alone does not stop it. With
        // setParseBigDecimal(true) it still hands back a Double for infinity/NaN, which the BigDecimal codec used to
        // cast blindly (ClassCastException); it is now rejected as a malformed value (issue M3 fix, adjacent finding).
        final byte[] bytes = stringSheet("G", new String[]{"BigDec"}, new String[][]{{"∞"}});   // U+221E, the Locale.ROOT infinity symbol
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "G", NumberPatternRow.class));
    }

    @Test
    public void collection_customPattern_validElements_parse() throws Exception {
        // A pattern on a Collection column applies to each element; well-formed elements bind as before.
        final byte[] bytes = stringSheet("C", new String[]{"Nums"}, new String[][]{{"10;1,200;30"}});
        assertThat(importList(bytes, "C", CollectionPatternRow.class).get(0).getNums()).containsExactly(10, 1200, 30);
    }

    @Test
    public void collection_customPattern_trailingGarbageElement_throws() throws Exception {
        // The element path goes through the same codec, so one bad element rejects the row rather than binding 2
        // for "2x" (issue M3 fix).
        final byte[] bytes = stringSheet("C", new String[]{"Nums"}, new String[][]{{"10;2x;30"}});
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "C", CollectionPatternRow.class));
    }

    @Test
    public void date_customPattern_trailingGarbage_throws() throws Exception {
        // java.util.Date had the same defect through SimpleDateFormat.parse (single-arg): "2024-01-02 xxx" parsed as
        // 2 January 2024. Now the custom pattern has to match in full, and so do the built-in read formatters it falls
        // back to (none of which is date-only), so the value is rejected (issue M2 fix).
        final byte[] bytes = stringSheet("R", new String[]{"D"}, new String[][]{{"2024-01-02 xxx"}});
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "R", LenientJavaDateRow.class));
    }

    @Test
    public void date_customPattern_exactValue_parses() throws Exception {
        // The value the pattern reads end to end still binds - the fallback chain is unchanged for well-formed input.
        final byte[] bytes = stringSheet("R", new String[]{"D"}, new String[][]{{"2024-01-02"}});

        final Calendar calendar = Calendar.getInstance();
        calendar.setTime(importList(bytes, "R", LenientJavaDateRow.class).get(0).getD());
        assertThat(calendar.get(Calendar.YEAR)).isEqualTo(2024);
        assertThat(calendar.get(Calendar.MONTH)).isEqualTo(Calendar.JANUARY);
        assertThat(calendar.get(Calendar.DAY_OF_MONTH)).isEqualTo(2);
    }

    // ==================================================================
    // Primitive codec edge paths (STRING cell, BOOLEAN cell, DecimalFormat pattern)
    // ==================================================================

    @Test
    public void primitiveIntegers_stringCell_parsed() throws Exception {
        // Reading each integer primitive from a STRING-typed cell exercises the codec's string-parse path.
        final byte[] bytes = stringSheet("P",
                new String[]{"PrimByte", "PrimShort", "PrimInt", "PrimLong"},
                new String[][]{{"12", "300", "70000", "9007199254740991"}});

        final AllTypesRow row = importList(bytes, "P", AllTypesRow.class).get(0);
        assertThat(row.getPrimByte()).isEqualTo((byte) 12);
        assertThat(row.getPrimShort()).isEqualTo((short) 300);
        assertThat(row.getPrimInt()).isEqualTo(70000);
        assertThat(row.getPrimLong()).isEqualTo(9007199254740991L);
    }

    @Test
    public void primitiveFloats_stringCell_parsed() throws Exception {
        final byte[] bytes = stringSheet("P",
                new String[]{"PrimFloat", "PrimDouble"},
                new String[][]{{"1.5", "2.25"}});

        final AllTypesRow row = importList(bytes, "P", AllTypesRow.class).get(0);
        assertThat(row.getPrimFloat()).isEqualTo(1.5F);
        assertThat(row.getPrimDouble()).isEqualTo(2.25);
    }

    @Test
    public void primitiveIntegers_blankStringCell_defaultsToZero() throws Exception {
        // A blank string cell parses to 0 for each integer primitive (the isBlank -> 0 branch).
        // A non-blank Text column keeps the row from being treated as a blank (skipped) row.
        final byte[] bytes = stringSheet("P",
                new String[]{"Text", "PrimByte", "PrimShort", "PrimInt", "PrimLong"},
                new String[][]{{"keep", "", "", "", ""}});

        final AllTypesRow row = importList(bytes, "P", AllTypesRow.class).get(0);
        assertThat(row.getText()).isEqualTo("keep");
        assertThat(row.getPrimByte()).isEqualTo((byte) 0);
        assertThat(row.getPrimShort()).isEqualTo((short) 0);
        assertThat(row.getPrimInt()).isEqualTo(0);
        assertThat(row.getPrimLong()).isEqualTo(0L);
    }

    @Test
    public void primitiveIntegers_booleanCell_mapToOneOrZero() throws Exception {
        // A BOOLEAN cell maps to 1 (true) / 0 (false) for integer primitives.
        final byte[] bytes = sheet("P", s -> {
            final Row header = s.createRow(0);
            header.createCell(0).setCellValue("PrimByte");
            header.createCell(1).setCellValue("PrimShort");
            header.createCell(2).setCellValue("PrimInt");
            header.createCell(3).setCellValue("PrimLong");
            final Row data = s.createRow(1);
            data.createCell(0).setCellValue(true);
            data.createCell(1).setCellValue(true);
            data.createCell(2).setCellValue(false);
            data.createCell(3).setCellValue(true);
        });

        final AllTypesRow row = importList(bytes, "P", AllTypesRow.class).get(0);
        assertThat(row.getPrimByte()).isEqualTo((byte) 1);
        assertThat(row.getPrimShort()).isEqualTo((short) 1);
        assertThat(row.getPrimInt()).isEqualTo(0);
        assertThat(row.getPrimLong()).isEqualTo(1L);
    }

    @Test
    public void numericWrappers_booleanCell_mapToOneOrZero() throws Exception {
        // A BOOLEAN cell maps to 1 (true) / 0 (false) for the wrapper integer types and both float types,
        // exercising the BOOLEAN branch of each of those codecs (the primitive integers are covered above).
        final byte[] bytes = sheet("P", s -> {
            final Row header = s.createRow(0);
            header.createCell(0).setCellValue("WrapByte");
            header.createCell(1).setCellValue("WrapShort");
            header.createCell(2).setCellValue("WrapInt");
            header.createCell(3).setCellValue("WrapLong");
            header.createCell(4).setCellValue("WrapDouble");
            header.createCell(5).setCellValue("WrapFloat");
            header.createCell(6).setCellValue("PrimDouble");
            header.createCell(7).setCellValue("PrimFloat");
            final Row data = s.createRow(1);
            data.createCell(0).setCellValue(true);
            data.createCell(1).setCellValue(false);
            data.createCell(2).setCellValue(true);
            data.createCell(3).setCellValue(false);
            data.createCell(4).setCellValue(true);
            data.createCell(5).setCellValue(false);
            data.createCell(6).setCellValue(true);
            data.createCell(7).setCellValue(false);
        });

        final AllTypesRow row = importList(bytes, "P", AllTypesRow.class).get(0);
        assertThat(row.getWrapByte()).isEqualTo((byte) 1);
        assertThat(row.getWrapShort()).isEqualTo((short) 0);
        assertThat(row.getWrapInt()).isEqualTo(1);
        assertThat(row.getWrapLong()).isEqualTo(0L);
        assertThat(row.getWrapDouble()).isEqualTo(1.0);
        assertThat(row.getWrapFloat()).isEqualTo(0.0F);
        assertThat(row.getPrimDouble()).isEqualTo(1.0);
        assertThat(row.getPrimFloat()).isEqualTo(0.0F);
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void primitives_decimalFormatPattern_roundTrip(final ExportDest dest) throws Exception {
        // A DecimalFormat pattern makes each primitive exported as text and re-parsed via DecimalFormat on import,
        // covering the primitive codec's exported-to-string / DecimalFormat branches on both directions.
        final PrimitivePatternRow row = new PrimitivePatternRow();
        row.setLongCount(1234567L);
        row.setIntCount(89012);
        row.setDoubleAmt(1234.5);

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(PrimitivePatternRow.class, Arrays.asList(row), "PP")
                .override(noValidationOption()), dest, testInfo);

        final PrimitivePatternRow out = pxl.importExcel()
                .sheet(PrimitivePatternRow.class, Arrays.asList("PP"))
                .fromStream(new ByteArrayInputStream(bytes)).get(0);
        assertThat(out.getLongCount()).isEqualTo(1234567L);
        assertThat(out.getIntCount()).isEqualTo(89012);
        assertThat(out.getDoubleAmt()).isEqualTo(1234.5);
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void numberWrappers_decimalFormatPattern_roundTrip(final ExportDest dest) throws Exception {
        // A DecimalFormat pattern on the wrapper (and remaining primitive) numeric types makes each value
        // exported as text and re-parsed via DecimalFormat on import, covering each numeric codec's
        // exported-to-string / DecimalFormat branch on both directions (chosen values survive the pattern round-trip exactly).
        final NumberPatternRow row = new NumberPatternRow();
        row.setWrapByte((byte) 100);
        row.setPrimByte((byte) -50);
        row.setWrapShort((short) 12345);
        row.setPrimShort((short) -6000);
        row.setWrapInt(1234567);
        row.setPrimInt(-7654321);
        row.setWrapLong(9876543L);
        row.setPrimLong(-1234567890L);
        row.setWrapFloat(-12.5F);
        row.setPrimFloat(3.5F);
        row.setWrapDouble(1234.25);
        row.setPrimDouble(-4321.75);
        row.setBigInt(new BigInteger("123456789"));
        row.setBigDec(new BigDecimal("98765.25"));

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(NumberPatternRow.class, Arrays.asList(row), "NP")
                .override(noValidationOption()), dest, testInfo);

        final NumberPatternRow out = pxl.importExcel()
                .sheet(NumberPatternRow.class, Arrays.asList("NP"))
                .fromStream(new ByteArrayInputStream(bytes)).get(0);
        assertThat(out.getWrapByte()).isEqualTo((byte) 100);
        assertThat(out.getPrimByte()).isEqualTo((byte) -50);
        assertThat(out.getWrapShort()).isEqualTo((short) 12345);
        assertThat(out.getPrimShort()).isEqualTo((short) -6000);
        assertThat(out.getWrapInt()).isEqualTo(1234567);
        assertThat(out.getPrimInt()).isEqualTo(-7654321);
        assertThat(out.getWrapLong()).isEqualTo(9876543L);
        assertThat(out.getPrimLong()).isEqualTo(-1234567890L);
        assertThat(out.getWrapFloat()).isEqualTo(-12.5F);
        assertThat(out.getPrimFloat()).isEqualTo(3.5F);
        assertThat(out.getWrapDouble()).isEqualTo(1234.25);
        assertThat(out.getPrimDouble()).isEqualTo(-4321.75);
        assertThat(out.getBigInt()).isEqualTo(new BigInteger("123456789"));
        assertThat(out.getBigDec()).isEqualByComparingTo(new BigDecimal("98765.25"));
    }

    // ==================================================================
    // Collection element-type coverage: round-trips a collection per element type,
    // exercising every element branch of the Collection codec plus each element codec's string path.
    // ==================================================================

    private CollectionTypesRow roundTripCollections(final ExportDest dest, final CollectionTypesRow row) throws Exception {
        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(CollectionTypesRow.class, Arrays.asList(row), "C")
                .override(noValidationOption()), dest, testInfo);
        return pxl.importExcel()
                .sheet(CollectionTypesRow.class, Arrays.asList("C"))
                .fromStream(new ByteArrayInputStream(bytes)).get(0);
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void collectionTypes_scalarElements_roundTrip(final ExportDest dest) throws Exception {
        final CollectionTypesRow row = new CollectionTypesRow();
        row.setBytes(Arrays.asList((byte) 1, (byte) -2));
        row.setShorts(Arrays.asList((short) 300, (short) -400));
        row.setLongs(Arrays.asList(10L, 20L));
        row.setDoubles(Arrays.asList(1.5, -2.25));
        row.setFloats(Arrays.asList(1.5F, -2.25F));
        row.setChars(Arrays.asList('A', 'z'));
        row.setBools(Arrays.asList(true, false));
        row.setBigInts(Arrays.asList(new BigInteger("123"), new BigInteger("-456")));
        row.setBigDecs(Arrays.asList(new BigDecimal("1.25"), new BigDecimal("-3.5")));
        row.setUuids(Arrays.asList(UUID.fromString("123e4567-e89b-12d3-a456-426614174000"),
                UUID.fromString("00112233-4455-6677-8899-aabbccddeeff")));

        final CollectionTypesRow out = roundTripCollections(dest, row);

        assertThat(out.getBytes()).containsExactly((byte) 1, (byte) -2);
        assertThat(out.getShorts()).containsExactly((short) 300, (short) -400);
        assertThat(out.getLongs()).containsExactly(10L, 20L);
        assertThat(out.getDoubles()).containsExactly(1.5, -2.25);
        assertThat(out.getFloats()).containsExactly(1.5F, -2.25F);
        assertThat(out.getChars()).containsExactly('A', 'z');
        assertThat(out.getBools()).containsExactly(true, false);
        assertThat(out.getBigInts()).containsExactly(new BigInteger("123"), new BigInteger("-456"));
        assertThat(out.getBigDecs()).hasSize(2);
        assertThat(out.getBigDecs().get(0)).isEqualByComparingTo("1.25");
        assertThat(out.getBigDecs().get(1)).isEqualByComparingTo("-3.5");
        assertThat(out.getUuids()).containsExactly(UUID.fromString("123e4567-e89b-12d3-a456-426614174000"),
                UUID.fromString("00112233-4455-6677-8899-aabbccddeeff"));
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void collectionTypes_dateTimeElements_roundTrip(final ExportDest dest) throws Exception {
        final ZoneId zone = ZoneId.systemDefault();
        final Date date1 = Date.from(LocalDateTime.of(2023, 1, 2, 3, 4, 5).atZone(zone).toInstant());
        final Date date2 = Date.from(LocalDateTime.of(2024, 6, 7, 8, 9, 10).atZone(zone).toInstant());

        final ZonedDateTime zdt1 = LocalDateTime.of(2023, 1, 2, 3, 4, 5).atZone(zone);
        final ZonedDateTime zdt2 = LocalDateTime.of(2024, 6, 7, 8, 9, 10).atZone(zone);

        final CollectionTypesRow row = new CollectionTypesRow();
        row.setLocalDates(Arrays.asList(LocalDate.of(2023, 1, 2), LocalDate.of(2024, 3, 4)));
        row.setLocalTimes(Arrays.asList(LocalTime.of(1, 2, 3), LocalTime.of(4, 5, 6)));
        row.setLocalDateTimes(Arrays.asList(LocalDateTime.of(2023, 1, 2, 3, 4, 5), LocalDateTime.of(2024, 6, 7, 8, 9, 10)));
        row.setZonedDateTimes(Arrays.asList(zdt1, zdt2));
        row.setOffsetTimes(Arrays.asList(zdt1.toOffsetDateTime().toOffsetTime(), zdt2.toOffsetDateTime().toOffsetTime()));
        row.setOffsetDateTimes(Arrays.asList(zdt1.toOffsetDateTime(), zdt2.toOffsetDateTime()));
        row.setJavaDates(Arrays.asList(date1, date2));
        row.setDurations(Arrays.asList(Duration.ofSeconds(1), Duration.ofHours(2).plusMinutes(3)));
        row.setPeriods(Arrays.asList(Period.of(1, 2, 3), Period.ofDays(5)));

        final CollectionTypesRow out = roundTripCollections(dest, row);

        assertThat(out.getLocalDates()).containsExactly(LocalDate.of(2023, 1, 2), LocalDate.of(2024, 3, 4));
        assertThat(out.getLocalTimes()).containsExactly(LocalTime.of(1, 2, 3), LocalTime.of(4, 5, 6));
        assertThat(out.getLocalDateTimes()).containsExactly(LocalDateTime.of(2023, 1, 2, 3, 4, 5), LocalDateTime.of(2024, 6, 7, 8, 9, 10));
        assertThat(out.getZonedDateTimes()).containsExactly(zdt1, zdt2);
        assertThat(out.getOffsetTimes()).containsExactly(zdt1.toOffsetDateTime().toOffsetTime(), zdt2.toOffsetDateTime().toOffsetTime());
        assertThat(out.getOffsetDateTimes()).containsExactly(zdt1.toOffsetDateTime(), zdt2.toOffsetDateTime());
        assertThat(out.getJavaDates()).containsExactly(date1, date2);
        assertThat(out.getDurations()).containsExactly(Duration.ofSeconds(1), Duration.ofHours(2).plusMinutes(3));
        assertThat(out.getPeriods()).containsExactly(Period.of(1, 2, 3), Period.ofDays(5));
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void collectionTypes_customObjectElements_roundTrip(final ExportDest dest) throws Exception {
        // A collection of a custom-convertible element type routes each element through the object codec.
        final CollectionTypesRow row = new CollectionTypesRow();
        row.setMoneys(Arrays.asList(new Money("USD", 100L), new Money("EUR", 200L)));

        final CollectionTypesRow out = roundTripCollections(dest, row);

        assertThat(out.getMoneys()).extracting(Money::getCurrency).containsExactly("USD", "EUR");
        assertThat(out.getMoneys()).extracting(Money::getAmount).containsExactly(100L, 200L);
    }

    // ==================================================================
    // CSV import of every supported type: every column arrives as a STRING, so this drives the
    // string dispatcher (PxlCellResolver#parseDataValueFromString) and each codec's string-parse path
    // (which the Excel round-trip mostly bypasses by reading NUMERIC/typed cells). The values are the
    // exportSample strings from AllTypesRow, which are guaranteed parseable as their column types.
    // ==================================================================

    @Test
    public void allTypes_csvImport_parsesEveryColumnFromString() throws Exception {
        final String[] headers = {
                "Text", "LeadingZero", "PrimByte", "WrapByte", "PrimShort", "WrapShort", "PrimInt", "WrapInt", "PrimLong", "WrapLong",
                "PrimDouble", "WrapDouble", "PrimFloat", "WrapFloat", "PrimChar", "WrapChar", "PrimBool", "WrapBool",
                "BigInt", "BigDec", "JavaDate", "LocalDate", "LocalTime", "LocalDateTime", "ZonedDateTime", "OffsetTime", "OffsetDateTime",
                "Duration", "Period", "Uuid", "Grade", "Category", "Point", "Money", "StringList", "IntList", "GradeList"};
        final String[] values = {
                "Sample text", "007", "1", "2", "3", "4", "5", "6", "7", "8",
                "1.5", "2.5", "3.5", "4.5", "A", "B", "true", "false",
                "12345678901234567890", "12345.6789", "2023-06-15 10:30:45", "2023-06-15", "10:30:45", "2023-06-15 10:30:45",
                "2023-06-15T10:30:45+09:00", "10:30:45+09:00", "2023-06-15T10:30:45+09:00",
                "PT1H2M3S", "P1Y2M3D", "123e4567-e89b-12d3-a456-426614174000",
                "A", "Electronics", "\"3,7\"", "USD 1050", "Apple;Banana;Cherry", "10;20;30", "A;B;F"};

        final String csv = String.join(",", headers) + "\n" + String.join(",", values) + "\n";

        final AllTypesRow row = pxl.importCsv()
                .sheet(AllTypesRow.class)
                .fromStream("AllTypes", new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8))).get(0);

        // strings
        assertThat(row.getText()).isEqualTo("Sample text");
        assertThat(row.getLeadingZero()).isEqualTo("007");
        // integer primitives / wrappers
        assertThat(row.getPrimByte()).isEqualTo((byte) 1);
        assertThat(row.getWrapByte()).isEqualTo((byte) 2);
        assertThat(row.getPrimShort()).isEqualTo((short) 3);
        assertThat(row.getWrapShort()).isEqualTo((short) 4);
        assertThat(row.getPrimInt()).isEqualTo(5);
        assertThat(row.getWrapInt()).isEqualTo(6);
        assertThat(row.getPrimLong()).isEqualTo(7L);
        assertThat(row.getWrapLong()).isEqualTo(8L);
        // floating point
        assertThat(row.getPrimDouble()).isEqualTo(1.5);
        assertThat(row.getWrapDouble()).isEqualTo(2.5);
        assertThat(row.getPrimFloat()).isEqualTo(3.5F);
        assertThat(row.getWrapFloat()).isEqualTo(4.5F);
        // char / boolean
        assertThat(row.getPrimChar()).isEqualTo('A');
        assertThat(row.getWrapChar()).isEqualTo('B');
        assertThat(row.isPrimBool()).isTrue();
        assertThat(row.getWrapBool()).isFalse();
        // big numbers
        assertThat(row.getBigInt()).isEqualTo(new BigInteger("12345678901234567890"));
        assertThat(row.getBigDec()).isEqualByComparingTo(new BigDecimal("12345.6789"));
        // date / time
        assertThat(row.getJavaDate()).isNotNull();
        assertThat(row.getLocalDate()).isEqualTo(LocalDate.of(2023, 6, 15));
        assertThat(row.getLocalTime()).isEqualTo(LocalTime.of(10, 30, 45));
        assertThat(row.getLocalDateTime()).isEqualTo(LocalDateTime.of(2023, 6, 15, 10, 30, 45));
        assertThat(row.getZonedDateTime()).isEqualTo(ZonedDateTime.parse("2023-06-15T10:30:45+09:00"));
        assertThat(row.getOffsetTime()).isEqualTo(OffsetTime.parse("10:30:45+09:00"));
        assertThat(row.getOffsetDateTime()).isEqualTo(OffsetDateTime.parse("2023-06-15T10:30:45+09:00"));
        assertThat(row.getDuration()).isEqualTo(Duration.ofHours(1).plusMinutes(2).plusSeconds(3));
        assertThat(row.getPeriod()).isEqualTo(Period.of(1, 2, 3));
        // UUID
        assertThat(row.getUuid()).isEqualTo(UUID.fromString("123e4567-e89b-12d3-a456-426614174000"));
        // enum (name match and toString-label match)
        assertThat(row.getGrade()).isEqualTo(Grade.A);
        assertThat(row.getCategory()).isEqualTo(Category.ELECTRONICS);
        // custom objects
        assertThat(row.getPoint().getX()).isEqualTo(3);
        assertThat(row.getPoint().getY()).isEqualTo(7);
        assertThat(row.getMoney().getCurrency()).isEqualTo("USD");
        assertThat(row.getMoney().getAmount()).isEqualTo(1050L);
        // collections
        assertThat(row.getStringList()).containsExactly("Apple", "Banana", "Cherry");
        assertThat(row.getIntList()).containsExactly(10, 20, 30);
        assertThat(row.getGradeList()).containsExactly(Grade.A, Grade.B, Grade.F);
    }

    // ==================================================================
    // Date/time types with a custom DateTimeFormatter pattern: each value is exported as formatted text and
    // re-parsed with the same formatter on import, covering each date/time codec's export-formatter and
    // import-formatter (cached-pattern) branches on both directions.
    // ==================================================================

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void dateTimeTypes_customPattern_roundTrip(final ExportDest dest) throws Exception {
        final DateTimePatternRow row = new DateTimePatternRow();
        row.setLocalDate(LocalDate.of(2023, 6, 15));
        row.setLocalTime(LocalTime.of(10, 30, 45));
        row.setLocalDateTime(LocalDateTime.of(2023, 6, 15, 10, 30, 45));
        row.setOffsetTime(OffsetTime.of(10, 30, 45, 0, ZoneOffset.ofHours(9)));
        row.setOffsetDateTime(OffsetDateTime.of(2023, 6, 15, 10, 30, 45, 0, ZoneOffset.ofHours(9)));

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(DateTimePatternRow.class, Arrays.asList(row), "DT")
                .override(noValidationOption()), dest, testInfo);

        final DateTimePatternRow out = pxl.importExcel()
                .sheet(DateTimePatternRow.class, Arrays.asList("DT"))
                .fromStream(new ByteArrayInputStream(bytes)).get(0);
        assertThat(out.getLocalDate()).isEqualTo(LocalDate.of(2023, 6, 15));
        assertThat(out.getLocalTime()).isEqualTo(LocalTime.of(10, 30, 45));
        assertThat(out.getLocalDateTime()).isEqualTo(LocalDateTime.of(2023, 6, 15, 10, 30, 45));
        assertThat(out.getOffsetTime()).isEqualTo(OffsetTime.of(10, 30, 45, 0, ZoneOffset.ofHours(9)));
        assertThat(out.getOffsetDateTime()).isEqualTo(OffsetDateTime.of(2023, 6, 15, 10, 30, 45, 0, ZoneOffset.ofHours(9)));
    }

    @Test
    public void dateTimeTypes_booleanCell_throws() throws Exception {
        // A BOOLEAN cell is no longer a defined conversion for any date/time type: each date/time codec
        // rejects it with PxlCellCodecException (it previously mapped to an Excel serial). Each type is
        // isolated in its own single-column sheet so every codec's BOOLEAN branch is exercised.
        final String[] headers = {"JavaDate", "LocalDate", "LocalTime", "LocalDateTime", "ZonedDateTime", "OffsetTime", "OffsetDateTime"};
        for (final String header : headers) {
            final byte[] bytes = sheet("B", s -> {
                s.createRow(0).createCell(0).setCellValue(header);
                s.createRow(1).createCell(0).setCellValue(true);
            });
            assertThrows(PxlCellCodecException.class,
                    () -> importList(bytes, "B", AllTypesRow.class),
                    header + " column: a BOOLEAN cell must be rejected with PxlCellCodecException");
        }
    }

    @Test
    public void dateTimeTypes_plainNumericCell_readAsExcelSerial() throws Exception {
        // A plain (non-date-formatted) NUMERIC cell is read as an Excel date serial for every date/time type,
        // exercising the "not date-formatted" numeric branch of each date/time codec.
        final byte[] bytes = sheet("N", s -> {
            final Row header = s.createRow(0);
            final String[] names = {"JavaDate", "LocalDate", "LocalTime", "LocalDateTime", "ZonedDateTime", "OffsetTime", "OffsetDateTime"};
            for (int i = 0; i < names.length; i++) {
                header.createCell(i).setCellValue(names[i]);
            }
            final Row data = s.createRow(1);
            for (int i = 0; i < names.length; i++) {
                data.createCell(i).setCellValue(45000.5);   // a plain (non-date-formatted) Excel serial
            }
        });

        final AllTypesRow row = importList(bytes, "N", AllTypesRow.class).get(0);
        assertThat(row.getJavaDate()).isNotNull();
        assertThat(row.getLocalDate()).isNotNull();
        assertThat(row.getLocalTime()).isNotNull();
        assertThat(row.getLocalDateTime()).isNotNull();
        assertThat(row.getZonedDateTime()).isNotNull();
        assertThat(row.getOffsetTime()).isNotNull();
        assertThat(row.getOffsetDateTime()).isNotNull();
    }

    @Test
    public void dateTimeTypes_noPattern_isoString_parsedViaIsoFallback() throws Exception {
        // With no column pattern, date/time strings parse via the fixed, locale-independent ISO read patterns
        // (or each codec's final ISO parser for the offset/zoned 'T' forms: ISO_OFFSET_* / ISO_ZONED for the
        // offset/zoned types). java.util.Date has no zone of its own, so its ISO 'T' read pattern
        // (SimpleDateFormat) resolves in the JVM default zone - the expected Date below is built with
        // ZoneId.systemDefault() so the assertion stays zone-independent. DateTimeNumericRow carries no pattern
        // on any column, and the result does not depend on the JVM default locale.
        final byte[] bytes = stringSheet("ISO",
                new String[]{"JavaDate", "LocalDate", "LocalTime", "LocalDateTime", "ZonedDateTime", "OffsetTime", "OffsetDateTime"},
                new String[][]{{
                        "2023-06-15T10:30:45",          // JavaDate: ISO-8601 date-time (no zone; read in the local zone)
                        "2023-06-15",                   // LocalDate: ISO-8601
                        "10:30:45",                     // LocalTime: ISO-8601
                        "2023-06-15T10:30:45",          // LocalDateTime: ISO-8601
                        "2023-06-15T10:30:45+09:00",    // ZonedDateTime: ISO-8601 with offset
                        "10:30:45+09:00",               // OffsetTime: ISO-8601 with offset
                        "2023-06-15T10:30:45+09:00"}}); // OffsetDateTime: ISO-8601 with offset

        final DateTimeNumericRow row = importList(bytes, "ISO", DateTimeNumericRow.class).get(0);
        // SimpleDateFormat parses at the JVM default zone, so build the expected Date the same way.
        assertThat(row.getJavaDate()).isEqualTo(Date.from(
                LocalDateTime.of(2023, 6, 15, 10, 30, 45).atZone(ZoneId.systemDefault()).toInstant()));
        assertThat(row.getLocalDate()).isEqualTo(LocalDate.of(2023, 6, 15));
        assertThat(row.getLocalTime()).isEqualTo(LocalTime.of(10, 30, 45));
        assertThat(row.getLocalDateTime()).isEqualTo(LocalDateTime.of(2023, 6, 15, 10, 30, 45));
        assertThat(row.getZonedDateTime()).isEqualTo(ZonedDateTime.parse("2023-06-15T10:30:45+09:00"));
        assertThat(row.getOffsetTime()).isEqualTo(OffsetTime.parse("10:30:45+09:00"));
        assertThat(row.getOffsetDateTime()).isEqualTo(OffsetDateTime.parse("2023-06-15T10:30:45+09:00"));
    }

    @Test
    public void dateTimeTypes_noPattern_legacySpaceSeparatedString_parsedViaReadFallback() throws Exception {
        // Backward compatibility: before the ISO 'T' switch the default write pattern used a space separator
        // (yyyy-MM-dd HH:mm:ss). The read patterns still accept that space form alongside the 'T' form, so a
        // date-time value written by an older PXL version round-trips unchanged. Locale-independent.
        final byte[] bytes = stringSheet("Legacy",
                new String[]{"JavaDate", "LocalDateTime"},
                new String[][]{{
                        "2023-06-15 10:30:45",          // JavaDate: legacy space separator
                        "2023-06-15 10:30:45"}});       // LocalDateTime: legacy space separator

        final DateTimeNumericRow row = importList(bytes, "Legacy", DateTimeNumericRow.class).get(0);
        assertThat(row.getLocalDateTime()).isEqualTo(LocalDateTime.of(2023, 6, 15, 10, 30, 45));
        // SimpleDateFormat parses at the JVM default zone, so build the expected Date the same way.
        assertThat(row.getJavaDate()).isEqualTo(Date.from(
                LocalDateTime.of(2023, 6, 15, 10, 30, 45).atZone(ZoneId.systemDefault()).toInstant()));
    }

    @Test
    public void offsetZonedTypes_noPattern_offsetlessString_throws() throws Exception {
        // Removed behavior: an offset/zone-less datetime string is no longer coerced with the system zone/offset.
        // With no pattern, Offset/Zoned columns parse strings only via the ISO-8601 formatter, which requires an
        // explicit offset/zone, so a bare (offset-less) value is rejected. Each type is isolated in its own sheet.
        final String[] dateTimeHeaders = {"ZonedDateTime", "OffsetDateTime"};
        for (final String header : dateTimeHeaders) {
            final byte[] bytes = stringSheet("NoOff", new String[]{header}, new String[][]{{"2023-06-15T10:30:45"}});
            assertThrows(PxlCellCodecException.class,
                    () -> importList(bytes, "NoOff", DateTimeNumericRow.class),
                    header + " column: an offset-less string must be rejected with PxlCellCodecException");
        }
        // OffsetTime: a bare time with no offset is likewise rejected.
        final byte[] timeBytes = stringSheet("NoOff", new String[]{"OffsetTime"}, new String[][]{{"10:30:45"}});
        assertThrows(PxlCellCodecException.class,
                () -> importList(timeBytes, "NoOff", DateTimeNumericRow.class),
                "OffsetTime column: an offset-less time must be rejected with PxlCellCodecException");
    }

    @Test
    public void dateTimeTypes_noPattern_localeFormattedString_throws() throws Exception {
        // Removed behavior: the JVM-locale read fallback is gone (full determinism). A no-pattern column parses
        // only the fixed ISO patterns, so a locale-formatted date such as the ko_KR-style "2023. 6. 15." is
        // rejected on every machine (it matches neither "y-M-d" nor ISO). Use a column pattern to accept it.
        final byte[] bytes = stringSheet("L", new String[]{"LocalDate"}, new String[][]{{"2023. 6. 15."}});
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "L", DateTimeNumericRow.class));
    }

    // ==================================================================
    // Enum codec: non-string cells, name() fallback, and custom converters
    // ==================================================================

    @Test
    public void enum_numericBooleanErrorCell_rejected() throws Exception {
        // A NUMERIC cell is stringified then matched (no enum matches -> parse error); a BOOLEAN cell and an
        // ERROR cell are both rejected as unsupported cell types.
        assertThrows(PxlCellCodecException.class, () -> importList(sheet("G", s -> {
            s.createRow(0).createCell(0).setCellValue("Grade");
            s.createRow(1).createCell(0).setCellValue(42);
        }), "G", AllTypesRow.class));

        assertThrows(PxlCellCodecException.class, () -> importList(sheet("G", s -> {
            s.createRow(0).createCell(0).setCellValue("Grade");
            s.createRow(1).createCell(0).setCellValue(true);
        }), "G", AllTypesRow.class));

        assertThrows(PxlCellCodecException.class, () -> importList(sheet("G", s -> {
            s.createRow(0).createCell(0).setCellValue("Grade");
            s.createRow(1).createCell(0).setCellErrorValue(FormulaError.DIV0.getCode());
        }), "G", AllTypesRow.class));
    }

    @Test
    public void enum_nameFallback_matchesByConstantName() throws Exception {
        // Category.FOOD's toString is "Food & Beverage"; importing its constant name "FOOD" matches via the
        // name() fallback after the toString match fails.
        final byte[] bytes = stringSheet("C", new String[]{"Category"}, new String[][]{{"FOOD"}});
        final AllTypesRow row = importList(bytes, "C", AllTypesRow.class).get(0);
        assertThat(row.getCategory()).isEqualTo(Category.FOOD);
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void enum_customConverter_roundTrip(final ExportDest dest) throws Exception {
        // ConverterEnum exports via @PxlExportConverter (toCode -> "2") and imports via @PxlImportConverter (fromCode).
        final ConverterEnumRow row = new ConverterEnumRow();
        row.setCode(ConverterEnum.TWO);

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(ConverterEnumRow.class, Arrays.asList(row), "E")
                .override(noValidationOption()), dest, testInfo);

        final ConverterEnumRow out = pxl.importExcel()
                .sheet(ConverterEnumRow.class, Arrays.asList("E"))
                .fromStream(new ByteArrayInputStream(bytes)).get(0);
        assertThat(out.getCode()).isEqualTo(ConverterEnum.TWO);
    }

    // ==================================================================
    // Object codec: a custom object whose export converter is a STATIC method
    // ==================================================================

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void object_staticExportConverter_roundTrip(final ExportDest dest) throws Exception {
        // StaticConverterObject exports via a static @PxlExportConverter (toStaticString) and imports via a
        // static @PxlImportConverter (fromString).
        final StaticConverterObjectRow row = new StaticConverterObjectRow();
        row.setValue(StaticConverterObject.fromString("hello"));

        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(StaticConverterObjectRow.class, Arrays.asList(row), "O")
                .override(noValidationOption()), dest, testInfo);

        final StaticConverterObjectRow out = pxl.importExcel()
                .sheet(StaticConverterObjectRow.class, Arrays.asList("O"))
                .fromStream(new ByteArrayInputStream(bytes)).get(0);
        assertThat(out.getValue().getValue()).isEqualTo("hello");
    }

    // ==================================================================
    // Numeric masking: exportMasking (no pattern) renders each numeric value as text and masks it
    // ==================================================================

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void numberMasking_exportMasksAllDigits(final ExportDest dest) throws Exception {
        final NumberMaskingRow row = new NumberMaskingRow();
        row.setWrapByte((byte) 12);
        row.setPrimByte((byte) 34);
        row.setWrapShort((short) 123);
        row.setPrimShort((short) 45);
        row.setWrapInt(6789);
        row.setPrimInt(12);
        row.setWrapLong(345L);
        row.setPrimLong(67L);
        row.setWrapDouble(12.5);
        row.setPrimDouble(3.5);
        row.setWrapFloat(4.5F);
        row.setPrimFloat(6.5F);
        row.setBigInt(new BigInteger("999"));
        row.setBigDec(new BigDecimal("12.34"));

        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(NumberMaskingRow.class, Arrays.asList(row), "M")
                .override(noValidationOption()), dest, testInfo)) {
            final Sheet sheet = workbook.getSheet("M");
            final Row header = sheet.getRow(0);
            final Row data = sheet.getRow(1);
            for (final Cell headerCell : header) {
                final Cell dataCell = data.getCell(headerCell.getColumnIndex());
                // the "\\d" mask replaces every digit with '*', so no digit remains in any numeric column
                assertThat(dataCell.getStringCellValue()).as(headerCell.getStringCellValue())
                        .isNotEmpty().doesNotContainPattern("[0-9]");
            }
        }
    }

    // ==================================================================
    // char/Character/Boolean masking and trim: the three codecs that render their own text now pass it through the
    // same string-level export processing as every other type, so exportMasking and exportTrim reach them too.
    // ==================================================================

    private static CharBoolMaskTrimRow charBoolMaskTrimRow() {
        // Every column is filled: a primitive char left at its default is taken as an absent value and would be
        // rendered as the export-null string, never reaching the masking or trimming under test.
        final CharBoolMaskTrimRow row = new CharBoolMaskTrimRow();
        row.setMaskWrapChar('x');
        row.setMaskPrimChar('y');
        row.setMaskBool(Boolean.TRUE);
        row.setTrimWrapChar(' ');
        row.setTrimPrimChar(' ');
        row.setTrimBool(Boolean.TRUE);
        return row;
    }

    // Locates the data cell by its header text; a value trimmed away leaves a blank cell, which POI may drop entirely.
    private static String charBoolCellOf(final Workbook workbook, final String headerName) {
        final Sheet sheet = workbook.getSheet("T");
        for (final Cell headerCell : sheet.getRow(0)) {
            if (headerName.equals(headerCell.getStringCellValue())) {
                final Cell dataCell = sheet.getRow(1).getCell(headerCell.getColumnIndex());
                return Objects.isNull(dataCell) ? "" : dataCell.getStringCellValue();
            }
        }
        return null;
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void charBoolMasking_exportMasksRenderedText(final ExportDest dest) throws Exception {
        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(CharBoolMaskTrimRow.class, Arrays.asList(charBoolMaskTrimRow()), "T")
                .override(noValidationOption()), dest, testInfo)) {
            // The "[a-z]" mask covers the single character as well as every letter of the rendered "true".
            assertThat(charBoolCellOf(workbook, "MaskWrapChar")).isEqualTo("*");
            assertThat(charBoolCellOf(workbook, "MaskPrimChar")).isEqualTo("*");
            assertThat(charBoolCellOf(workbook, "MaskBool")).isEqualTo("****");
        }
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void charBoolTrim_exportTrimsRenderedText(final ExportDest dest) throws Exception {
        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(CharBoolMaskTrimRow.class, Arrays.asList(charBoolMaskTrimRow()), "T")
                .override(noValidationOption()), dest, testInfo)) {
            // A whitespace character trims away entirely; the padded true string keeps only its letter.
            assertThat(charBoolCellOf(workbook, "TrimWrapChar")).isEmpty();
            assertThat(charBoolCellOf(workbook, "TrimPrimChar")).isEmpty();
            assertThat(charBoolCellOf(workbook, "TrimBool")).isEqualTo("Y");
        }
    }

    // ==================================================================
    // Unset primitive char: a char cannot be null, so a field that was never set holds (char) 0 - the type's only way
    // of saying "no value". Export takes it as absent and writes the column's export-null string. Writing the NUL
    // character itself would leave it to the XLSX writers, which silently replace it with '?' while saving, so the
    // round trip would slide from '\0' to '?' to the character '?'.
    // ==================================================================

    private static UnsetCharRow unsetCharRow() {
        final UnsetCharRow row = new UnsetCharRow();
        row.setSetChar('A');   // the control column, the only one given a value
        return row;
    }

    // Locates the data cell by its header text, so the assertion does not depend on the column order.
    private static String unsetCharCellOf(final Workbook workbook, final String headerName) {
        final Sheet sheet = workbook.getSheet("C");
        for (final Cell headerCell : sheet.getRow(0)) {
            if (headerName.equals(headerCell.getStringCellValue())) {
                final Cell dataCell = sheet.getRow(1).getCell(headerCell.getColumnIndex());
                return Objects.isNull(dataCell) ? null : dataCell.getStringCellValue();
            }
        }
        return null;
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void charExport_unsetPrimitiveChar_writesExportNullString(final ExportDest dest) throws Exception {
        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(UnsetCharRow.class, Arrays.asList(unsetCharRow()), "C")
                .override(noValidationOption()), dest, testInfo)) {
            // Asserts what PXL wrote rather than what a writer might have replaced: WORKBOOK holds the live workbook,
            // where a NUL would still be a NUL, so the three destinations agree only because no NUL was ever written.
            assertThat(unsetCharCellOf(workbook, "UnsetChar")).isEqualTo("");
            assertThat(unsetCharCellOf(workbook, "UnsetCharDash")).isEqualTo("-");
            assertThat(unsetCharCellOf(workbook, "SetChar")).isEqualTo("A");
            // The boxed column takes the same path through null, as it always has.
            assertThat(unsetCharCellOf(workbook, "UnsetWrapChar")).isEqualTo("");
        }
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void charExport_unsetPrimitiveChar_roundTripsAsUnset(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(UnsetCharRow.class, Arrays.asList(unsetCharRow()), "C")
                .override(noValidationOption()), dest, testInfo);

        final List<UnsetCharRow> out = importList(bytes, "C", UnsetCharRow.class);

        // The empty cell parses back to (char) 0, so the unset field comes back unset instead of holding '?'.
        assertThat(out.get(0).getUnsetChar()).isEqualTo((char) 0);
        assertThat(out.get(0).getSetChar()).isEqualTo('A');
        assertThat(out.get(0).getUnsetWrapChar()).isNull();
    }

    // ==================================================================
    // UUID: a value that only has meaning as text. Import accepts the canonical 8-4-4-4-12 form in either case and
    // nothing else; export always writes that form in lower case. The strictness belongs to the codec rather than to
    // UUID.fromString, which counts the hyphen-separated groups but not their digits and so would widen "1-1-1-1-1"
    // into an entirely different value.
    // ==================================================================

    private static final String UUID_TEXT = "123e4567-e89b-12d3-a456-426614174000";

    private static final String OTHER_UUID_TEXT = "00112233-4455-6677-8899-aabbccddeeff";

    private static final UUID UUID_VALUE = UUID.fromString(UUID_TEXT);

    private static final UUID OTHER_UUID_VALUE = UUID.fromString(OTHER_UUID_TEXT);

    private UuidRow roundTripUuid(final ExportDest dest, final UuidRow row) throws Exception {
        final byte[] bytes = emit(pxl.exportExcel()
                .sheet(UuidRow.class, Arrays.asList(row), "U")
                .override(noValidationOption()), dest, testInfo);
        return pxl.importExcel()
                .sheet(UuidRow.class, Arrays.asList("U"))
                .fromStream(new ByteArrayInputStream(bytes)).get(0);
    }

    private Workbook exportUuidWorkbook(final ExportDest dest, final UuidRow row) throws Exception {
        return workbookOf(pxl.exportExcel()
                .sheet(UuidRow.class, Arrays.asList(row), "U")
                .override(noValidationOption()), dest, testInfo);
    }

    // Locates the data cell by its header text, so the assertion does not depend on the column order.
    private static String dataStringOf(final Workbook workbook, final String headerName) {
        final Sheet sheet = workbook.getSheet("U");
        for (final Cell headerCell : sheet.getRow(0)) {
            if (headerName.equals(headerCell.getStringCellValue())) {
                return sheet.getRow(1).getCell(headerCell.getColumnIndex()).getStringCellValue();
            }
        }
        return null;
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void uuid_roundTrip_preservesValue(final ExportDest dest) throws Exception {
        final UuidRow row = new UuidRow();
        row.setId(UUID_VALUE);
        row.setExact(OTHER_UUID_VALUE);

        final UuidRow out = roundTripUuid(dest, row);

        assertThat(out.getId()).isEqualTo(UUID_VALUE);
        assertThat(out.getExact()).isEqualTo(OTHER_UUID_VALUE);
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void uuid_nilValue_roundTrips(final ExportDest dest) throws Exception {
        // The nil UUID is an ordinary value, not a stand-in for null: folding it into null would break the round-trip.
        final UuidRow row = new UuidRow();
        row.setId(new UUID(0L, 0L));

        final UuidRow out = roundTripUuid(dest, row);

        assertThat(out.getId()).isEqualTo(new UUID(0L, 0L));
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void uuid_export_writesCanonicalLowerCaseText(final ExportDest dest) throws Exception {
        final UuidRow row = new UuidRow();
        row.setId(UUID.fromString(UUID_TEXT.toUpperCase(Locale.ROOT)));

        final Workbook workbook = exportUuidWorkbook(dest, row);
        try {
            assertThat(dataStringOf(workbook, "Id")).isEqualTo(UUID_TEXT);
        } finally {
            workbook.close();
        }
    }

    @Test
    public void uuid_upperCaseInput_bindsSameValue() throws Exception {
        final byte[] bytes = stringSheet("U", new String[]{"Id"}, new String[][]{{UUID_TEXT.toUpperCase(Locale.ROOT)}});
        assertThat(importList(bytes, "U", UuidRow.class).get(0).getId()).isEqualTo(UUID_VALUE);
    }

    @Test
    public void uuid_shortGroups_throws() throws Exception {
        // UUID.fromString accepts this and widens it into 00000001-0001-0001-0001-000000000001; the codec must not.
        final byte[] bytes = stringSheet("U", new String[]{"Id"}, new String[][]{{"1-1-1-1-1"}});
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "U", UuidRow.class));
    }

    @Test
    public void uuid_hyphenlessHex_throws() throws Exception {
        // Export only ever writes the canonical form, so import accepts only that form.
        final byte[] bytes = stringSheet("U", new String[]{"Id"}, new String[][]{{"123e4567e89b12d3a456426614174000"}});
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "U", UuidRow.class));
    }

    @Test
    public void uuid_bracedForm_throws() throws Exception {
        final byte[] bytes = stringSheet("U", new String[]{"Id"}, new String[][]{{"{" + UUID_TEXT + "}"}});
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "U", UuidRow.class));
    }

    @Test
    public void uuid_urnPrefixedForm_throws() throws Exception {
        final byte[] bytes = stringSheet("U", new String[]{"Id"}, new String[][]{{"urn:uuid:" + UUID_TEXT}});
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "U", UuidRow.class));
    }

    @Test
    public void uuid_blankCell_bindsNull() throws Exception {
        final byte[] bytes = sheet("U", s -> {
            final Row header = s.createRow(0);
            header.createCell(0).setCellValue("Id");
            header.createCell(1).setCellValue("Exact");
            final Row data = s.createRow(1);
            data.createCell(0).setBlank();
            data.createCell(1).setCellValue(UUID_TEXT);
        });

        final UuidRow row = importList(bytes, "U", UuidRow.class).get(0);

        assertThat(row.getId()).isNull();
        assertThat(row.getExact()).isEqualTo(UUID_VALUE);
    }

    @Test
    public void uuid_booleanCell_throws() throws Exception {
        final byte[] bytes = sheet("U", s -> {
            s.createRow(0).createCell(0).setCellValue("Id");
            s.createRow(1).createCell(0).setCellValue(true);
        });
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "U", UuidRow.class));
    }

    @Test
    public void uuid_numericCell_throws() throws Exception {
        // A NUMERIC cell is rendered to text first, so the failure reports an invalid value rather than a cell type.
        final byte[] bytes = sheet("U", s -> {
            s.createRow(0).createCell(0).setCellValue("Id");
            s.createRow(1).createCell(0).setCellValue(123456);
        });
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "U", UuidRow.class));
    }

    @Test
    public void uuid_paddedValue_trimsAndBinds() throws Exception {
        final byte[] bytes = stringSheet("U", new String[]{"Id"}, new String[][]{{"  " + UUID_TEXT + "  "}});
        assertThat(importList(bytes, "U", UuidRow.class).get(0).getId()).isEqualTo(UUID_VALUE);
    }

    @Test
    public void uuid_untrimmedValue_throws() throws Exception {
        // The Exact column disables importTrim, so the padding stays part of the value and it is no longer canonical.
        final byte[] bytes = stringSheet("U", new String[]{"Exact"}, new String[][]{{" " + UUID_TEXT + " "}});
        assertThrows(PxlCellCodecException.class, () -> importList(bytes, "U", UuidRow.class));
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void uuid_exportMasking_masksHexDigits(final ExportDest dest) throws Exception {
        final UuidRow row = new UuidRow();
        row.setMasked(UUID_VALUE);

        final Workbook workbook = exportUuidWorkbook(dest, row);
        try {
            // The "[0-9a-f]" mask replaces every hexadecimal digit of the canonical form, leaving its hyphens.
            assertThat(dataStringOf(workbook, "Masked")).isEqualTo("********-****-****-****-************");
        } finally {
            workbook.close();
        }
    }

    @Test
    public void uuid_importUnique_duplicateThrows() throws Exception {
        // The same UUID written in two cases is one value - a String column would not have caught this pair.
        final byte[] bytes = stringSheet("U", new String[]{"Unique"},
                new String[][]{{UUID_TEXT}, {UUID_TEXT.toUpperCase(Locale.ROOT)}});
        assertThrows(PxlValidationException.class, () -> importList(bytes, "U", UuidRow.class));
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void uuidCollection_roundTrip_preservesNullPositions(final ExportDest dest) throws Exception {
        final UuidRow row = new UuidRow();
        row.setIds(Arrays.asList(UUID_VALUE, null, OTHER_UUID_VALUE));

        final UuidRow out = roundTripUuid(dest, row);

        assertThat(out.getIds()).containsExactly(UUID_VALUE, null, OTHER_UUID_VALUE);
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void uuidCollection_export_joinsCanonicalStrings(final ExportDest dest) throws Exception {
        // Guards the collection element branch of the dispatcher: a UUID element type resolves to no custom converter
        // any more, so without that branch this export fails as an unsupported element type.
        final UuidRow row = new UuidRow();
        row.setIds(Arrays.asList(UUID_VALUE, OTHER_UUID_VALUE));

        final Workbook workbook = exportUuidWorkbook(dest, row);
        try {
            assertThat(dataStringOf(workbook, "Ids")).isEqualTo(UUID_TEXT + ";" + OTHER_UUID_TEXT);
        } finally {
            workbook.close();
        }
    }
}
