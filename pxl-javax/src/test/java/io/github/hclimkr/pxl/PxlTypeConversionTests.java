package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.exception.PxlReflectionException;
import io.github.hclimkr.pxl.exception.PxlValidationException;
import io.github.hclimkr.pxl.tcdata.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInfo;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.time.*;
import java.util.*;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Per-type value conversion (codec) tests.
 * <p>
 * Round-trips each supported type through export→import with various values (boundary, negative, zero,
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
    private List<AllTypesRow> roundTrip(final List<AllTypesRow> rows) throws Exception {
        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("Types", rows, AllTypesRow.class)
                .override(noValidationOption())
                .toFile(excelFile);
        return pxl.importExcel()
                .sheet(AllTypesRow.class, Arrays.asList("Types"))
                .fromFile(excelFile);
    }

    private AllTypesRow roundTrip1(final AllTypesRow row) throws Exception {
        return roundTrip(Arrays.asList(row)).get(0);
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

    @Test
    public void integerTypes_boundaryValues_roundTrip() throws Exception {
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

        final List<AllTypesRow> out = roundTrip(Arrays.asList(max, min, zero));

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

    @Test
    public void integerWrappers_null_roundTripsAsNull() throws Exception {
        final AllTypesRow row = Fixtures.baseAllTypesRow();
        row.setWrapByte(null);
        row.setWrapShort(null);
        row.setWrapInt(null);
        row.setWrapLong(null);

        final AllTypesRow out = roundTrip1(row);

        assertThat(out.getWrapByte()).isNull();
        assertThat(out.getWrapShort()).isNull();
        assertThat(out.getWrapInt()).isNull();
        assertThat(out.getWrapLong()).isNull();
        assertThat(out.getPrimInt()).isEqualTo(row.getPrimInt());
    }

    @Test
    public void floatTypes_variousValues_roundTrip() throws Exception {
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

        final List<AllTypesRow> out = roundTrip(Arrays.asList(a, b, nulls));

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

    @Test
    public void bigNumbers_variousValues_roundTrip() throws Exception {
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

        final List<AllTypesRow> out = roundTrip(Arrays.asList(big, neg, zeroScale, nulls));

        assertThat(out.get(0).getBigInt()).isEqualTo(new BigInteger("123456789012345678901234567890"));
        assertThat(out.get(0).getBigDec()).isEqualByComparingTo(new BigDecimal("3.14159265358979323846"));
        assertThat(out.get(1).getBigInt()).isEqualTo(new BigInteger("-98765432109876543210"));
        assertThat(out.get(1).getBigDec()).isEqualByComparingTo(new BigDecimal("-0.0001"));
        assertThat(out.get(2).getBigInt()).isEqualTo(BigInteger.ZERO);
        assertThat(out.get(2).getBigDec()).isEqualTo(new BigDecimal("100.00"));   // scale (2) preserved
        assertThat(out.get(3).getBigInt()).isNull();
        assertThat(out.get(3).getBigDec()).isNull();
    }

    @Test
    public void boolean_trueFalseNull_roundTrip() throws Exception {
        final AllTypesRow t = Fixtures.baseAllTypesRow();
        t.setPrimBool(true);
        t.setWrapBool(Boolean.TRUE);
        final AllTypesRow f = Fixtures.baseAllTypesRow();
        f.setPrimBool(false);
        f.setWrapBool(Boolean.FALSE);
        final AllTypesRow n = Fixtures.baseAllTypesRow();
        n.setWrapBool(null);

        final List<AllTypesRow> out = roundTrip(Arrays.asList(t, f, n));

        assertThat(out.get(0).isPrimBool()).isTrue();
        assertThat(out.get(0).getWrapBool()).isTrue();
        assertThat(out.get(1).isPrimBool()).isFalse();
        assertThat(out.get(1).getWrapBool()).isFalse();
        assertThat(out.get(2).getWrapBool()).isNull();
    }

    @Test
    public void char_variousValues_roundTrip() throws Exception {
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

        final List<AllTypesRow> out = roundTrip(Arrays.asList(letter, digit, symbol, nullChar));

        assertThat(out.get(0).getPrimChar()).isEqualTo('Z');
        assertThat(out.get(0).getWrapChar()).isEqualTo('a');
        assertThat(out.get(1).getPrimChar()).isEqualTo('7');
        assertThat(out.get(1).getWrapChar()).isEqualTo('0');
        assertThat(out.get(2).getPrimChar()).isEqualTo('#');
        assertThat(out.get(2).getWrapChar()).isEqualTo('@');
        assertThat(out.get(3).getWrapChar()).isNull();
    }

    @Test
    public void string_specialCases_roundTrip() throws Exception {
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

        final List<AllTypesRow> out = roundTrip(Arrays.asList(
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

    @Test
    public void enum_allConstantsAndNull_roundTrip() throws Exception {
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

        final List<AllTypesRow> out = roundTrip(rows);

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

    @Test
    public void dateTime_variousValues_roundTrip() throws Exception {
        final ZoneId zone = ZoneId.systemDefault();

        final LocalDateTime midnight = LocalDateTime.of(2000, 1, 1, 0, 0, 0);
        final AllTypesRow a = Fixtures.baseAllTypesRow();
        a.setJavaDate(java.util.Date.from(midnight.atZone(zone).toInstant()));
        a.setLocalDate(LocalDate.of(2000, 1, 1));
        a.setLocalTime(LocalTime.of(0, 0, 0));
        a.setLocalDateTime(midnight);
        a.setZonedDateTime(midnight.atZone(zone));
        a.setOffsetDateTime(midnight.atZone(zone).toOffsetDateTime());
        a.setDuration(Duration.ZERO);
        a.setPeriod(Period.ZERO);

        final LocalDateTime endOfDay = LocalDateTime.of(2038, 12, 31, 23, 59, 59);
        final AllTypesRow b = Fixtures.baseAllTypesRow();
        b.setJavaDate(java.util.Date.from(endOfDay.atZone(zone).toInstant()));
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

        final List<AllTypesRow> out = roundTrip(Arrays.asList(a, b, nulls));

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

    @Test
    public void customObject_variousValues_roundTrip() throws Exception {
        final AllTypesRow a = Fixtures.baseAllTypesRow();
        a.setPoint(new Point(-3, -7));
        a.setMoney(new Money("EUR", 999999999L));
        final AllTypesRow b = Fixtures.baseAllTypesRow();
        b.setPoint(new Point(0, 0));
        b.setMoney(new Money("JPY", 0L));
        final AllTypesRow nulls = Fixtures.baseAllTypesRow();
        nulls.setPoint(null);
        nulls.setMoney(null);

        final List<AllTypesRow> out = roundTrip(Arrays.asList(a, b, nulls));

        assertThat(out.get(0).getPoint().getX()).isEqualTo(-3);
        assertThat(out.get(0).getPoint().getY()).isEqualTo(-7);
        assertThat(out.get(0).getMoney().getCurrency()).isEqualTo("EUR");
        assertThat(out.get(0).getMoney().getAmount()).isEqualTo(999999999L);
        assertThat(out.get(1).getPoint().getX()).isEqualTo(0);
        assertThat(out.get(1).getMoney().getAmount()).isEqualTo(0L);
        assertThat(out.get(2).getPoint()).isNull();
        assertThat(out.get(2).getMoney()).isNull();
    }

    @Test
    public void collection_positionFidelity_roundTrip() throws Exception {
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

        final List<AllTypesRow> out = roundTrip(Arrays.asList(multi, single, innerNull, edgeNull));

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

    @Test
    public void collection_emptyAndNull_roundTripsAsNull() throws Exception {
        final AllTypesRow empty = Fixtures.baseAllTypesRow();
        empty.setStringList(new ArrayList<>());
        empty.setIntList(new ArrayList<>());
        empty.setGradeList(new ArrayList<>());
        final AllTypesRow nulls = Fixtures.baseAllTypesRow();
        nulls.setStringList(null);
        nulls.setIntList(null);
        nulls.setGradeList(null);

        final List<AllTypesRow> out = roundTrip(Arrays.asList(empty, nulls));

        // An empty collection is exported as an empty cell and becomes null on import.
        assertThat(out.get(0).getStringList()).isNull();
        assertThat(out.get(0).getIntList()).isNull();
        assertThat(out.get(0).getGradeList()).isNull();
        assertThat(out.get(1).getStringList()).isNull();
    }

    @Test
    public void allNullableFields_null_roundTrip() throws Exception {
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

        final AllTypesRow out = roundTrip1(row);

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

    @Test
    public void long_withPattern_preservesLargeValue() throws Exception {
        final long big = 9007199254740993L;   // 2^53 + 1

        final LongPatternRow row = new LongPatternRow();
        row.setBig(big);

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("Big", Arrays.asList(row), LongPatternRow.class)
                .override(noValidationOption())
                .toFile(excelFile);

        final LongPatternRow out = pxl.importExcel()
                .sheet(LongPatternRow.class, Arrays.asList("Big"))
                .fromFile(excelFile).get(0);
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
            final org.apache.poi.xssf.usermodel.XSSFCell cell =
                    (org.apache.poi.xssf.usermodel.XSSFCell) s.createRow(1).createCell(0);
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
        // requireFinite has no range check - a finite NUMERIC value far beyond the Long range (~9.2e18) must also
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
    public void numeric_customPattern_trailingGarbage_characterize() throws Exception {
        // A numeric custom pattern uses DecimalFormat.parse (single-arg), so it silently drops trailing garbage ("abc" of "123abc") and parses 123 (issue M3, characterization).
        // Pins the asymmetry with the no-pattern case, which rejects via full validation.
        final byte[] bytes = stringSheet("G", new String[]{"Amount"}, new String[][]{{"123abc"}});
        assertThat(importList(bytes, "G", PatternRow.class).get(0).getAmount()).isEqualByComparingTo(new BigDecimal("123"));
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

    @Test
    public void primitives_decimalFormatPattern_roundTrip() throws Exception {
        // A DecimalFormat pattern makes each primitive exported as text and re-parsed via DecimalFormat on import,
        // covering the primitive codec's exported-to-string / DecimalFormat branches on both directions.
        final PrimitivePatternRow row = new PrimitivePatternRow();
        row.setLongCount(1234567L);
        row.setIntCount(89012);
        row.setDoubleAmt(1234.5);

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("PP", Arrays.asList(row), PrimitivePatternRow.class)
                .override(noValidationOption())
                .toFile(excelFile);

        final PrimitivePatternRow out = pxl.importExcel()
                .sheet(PrimitivePatternRow.class, Arrays.asList("PP"))
                .fromFile(excelFile).get(0);
        assertThat(out.getLongCount()).isEqualTo(1234567L);
        assertThat(out.getIntCount()).isEqualTo(89012);
        assertThat(out.getDoubleAmt()).isEqualTo(1234.5);
    }

    @Test
    public void numberWrappers_decimalFormatPattern_roundTrip() throws Exception {
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

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("NP", Arrays.asList(row), NumberPatternRow.class)
                .override(noValidationOption())
                .toFile(excelFile);

        final NumberPatternRow out = pxl.importExcel()
                .sheet(NumberPatternRow.class, Arrays.asList("NP"))
                .fromFile(excelFile).get(0);
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

    private CollectionTypesRow roundTripCollections(final CollectionTypesRow row) throws Exception {
        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("C", Arrays.asList(row), CollectionTypesRow.class)
                .override(noValidationOption())
                .toFile(excelFile);
        return pxl.importExcel()
                .sheet(CollectionTypesRow.class, Arrays.asList("C"))
                .fromFile(excelFile).get(0);
    }

    @Test
    public void collectionTypes_scalarElements_roundTrip() throws Exception {
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

        final CollectionTypesRow out = roundTripCollections(row);

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
    }

    @Test
    public void collectionTypes_dateTimeElements_roundTrip() throws Exception {
        final ZoneId zone = ZoneId.systemDefault();
        final java.util.Date date1 = java.util.Date.from(LocalDateTime.of(2023, 1, 2, 3, 4, 5).atZone(zone).toInstant());
        final java.util.Date date2 = java.util.Date.from(LocalDateTime.of(2024, 6, 7, 8, 9, 10).atZone(zone).toInstant());

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

        final CollectionTypesRow out = roundTripCollections(row);

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

    @Test
    public void collectionTypes_customObjectElements_roundTrip() throws Exception {
        // A collection of a custom-convertible element type routes each element through the object codec.
        final CollectionTypesRow row = new CollectionTypesRow();
        row.setMoneys(Arrays.asList(new Money("USD", 100L), new Money("EUR", 200L)));

        final CollectionTypesRow out = roundTripCollections(row);

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
                "Duration", "Period", "Grade", "Category", "Point", "Money", "StringList", "IntList", "GradeList"};
        final String[] values = {
                "Sample text", "007", "1", "2", "3", "4", "5", "6", "7", "8",
                "1.5", "2.5", "3.5", "4.5", "A", "B", "true", "false",
                "12345678901234567890", "12345.6789", "2023-06-15 10:30:45", "2023-06-15", "10:30:45", "2023-06-15 10:30:45",
                "2023-06-15T10:30:45+09:00", "10:30:45+09:00", "2023-06-15T10:30:45+09:00",
                "PT1H2M3S", "P1Y2M3D", "A", "Electronics", "\"3,7\"", "USD 1050", "Apple;Banana;Cherry", "10;20;30", "A;B;F"};

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

    @Test
    public void dateTimeTypes_customPattern_roundTrip() throws Exception {
        final DateTimePatternRow row = new DateTimePatternRow();
        row.setLocalDate(LocalDate.of(2023, 6, 15));
        row.setLocalTime(LocalTime.of(10, 30, 45));
        row.setLocalDateTime(LocalDateTime.of(2023, 6, 15, 10, 30, 45));
        row.setOffsetTime(OffsetTime.of(10, 30, 45, 0, ZoneOffset.ofHours(9)));
        row.setOffsetDateTime(OffsetDateTime.of(2023, 6, 15, 10, 30, 45, 0, ZoneOffset.ofHours(9)));

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("DT", Arrays.asList(row), DateTimePatternRow.class)
                .override(noValidationOption())
                .toFile(excelFile);

        final DateTimePatternRow out = pxl.importExcel()
                .sheet(DateTimePatternRow.class, Arrays.asList("DT"))
                .fromFile(excelFile).get(0);
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
        // (SimpleDateFormat) resolves in the JVM default zone — the expected Date below is built with
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

    @Test
    public void enum_customConverter_roundTrip() throws Exception {
        // ConverterEnum exports via @PxlExportConverter (toCode -> "2") and imports via @PxlImportConverter (fromCode).
        final ConverterEnumRow row = new ConverterEnumRow();
        row.setCode(ConverterEnum.TWO);

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("E", Arrays.asList(row), ConverterEnumRow.class)
                .override(noValidationOption())
                .toFile(excelFile);

        final ConverterEnumRow out = pxl.importExcel()
                .sheet(ConverterEnumRow.class, Arrays.asList("E"))
                .fromFile(excelFile).get(0);
        assertThat(out.getCode()).isEqualTo(ConverterEnum.TWO);
    }

    // ==================================================================
    // Object codec: a custom object whose export converter is a STATIC method
    // ==================================================================

    @Test
    public void object_staticExportConverter_roundTrip() throws Exception {
        // StaticConverterObject exports via a static @PxlExportConverter (toStaticString) and imports via a
        // static @PxlImportConverter (fromString).
        final StaticConverterObjectRow row = new StaticConverterObjectRow();
        row.setValue(StaticConverterObject.fromString("hello"));

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet("O", Arrays.asList(row), StaticConverterObjectRow.class)
                .override(noValidationOption())
                .toFile(excelFile);

        final StaticConverterObjectRow out = pxl.importExcel()
                .sheet(StaticConverterObjectRow.class, Arrays.asList("O"))
                .fromFile(excelFile).get(0);
        assertThat(out.getValue().getValue()).isEqualTo("hello");
    }

    // ==================================================================
    // Numeric masking: exportMasking (no pattern) renders each numeric value as text and masks it
    // ==================================================================

    @Test
    public void numberMasking_exportMasksAllDigits() throws Exception {
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

        final Workbook workbook = pxl.exportExcel()
                .sheet("M", Arrays.asList(row), NumberMaskingRow.class)
                .override(noValidationOption())
                .toWorkbook();
        try {
            final Sheet sheet = workbook.getSheet("M");
            final Row header = sheet.getRow(0);
            final Row data = sheet.getRow(1);
            for (final Cell headerCell : header) {
                final Cell dataCell = data.getCell(headerCell.getColumnIndex());
                // the "\\d" mask replaces every digit with '*', so no digit remains in any numeric column
                assertThat(dataCell.getStringCellValue()).as(headerCell.getStringCellValue())
                        .isNotEmpty().doesNotContainPattern("[0-9]");
            }
        } finally {
            workbook.close();
        }
    }
}
