package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.*;
import java.util.Arrays;
import java.util.Date;
import java.util.UUID;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Standard data builders and verification helpers for tests.
 * <p>
 * All data is written in English (ASCII). Dates/times use only down to the second so they can round-trip with the default pattern,
 * and the Zoned/Offset types are built with the system default zone/offset so they are restored equivalently even after the default export (which loses the zone/offset).
 */
public final class Fixtures {

    private Fixtures() {
    }

    /**
     * An export option with data validation turned off, to round-trip with a minimal file and no enum dropdown (hidden sheet).
     */
    public static PxlExportWorkbookOption noValidationOption() {
        return PxlExportWorkbookOption.builder().exportDataValidation(false).build();
    }

    // The reference date/time for round-trip verification (second precision, no nanoseconds)
    public static final LocalDateTime BASE_DATE_TIME = LocalDateTime.of(2023, 6, 15, 10, 30, 45);

    /**
     * Creates a standard AllTypesRow instance populated with every supported type.
     */
    public static AllTypesRow sampleAllTypesRow() {

        final ZoneId zone = ZoneId.systemDefault();

        final AllTypesRow row = new AllTypesRow();

        // String
        row.setText("Hello, PXL");
        row.setLeadingZero("007");

        // Integer types (including boundary values)
        row.setPrimByte((byte) -128);
        row.setWrapByte((byte) 127);
        row.setPrimShort((short) -32768);
        row.setWrapShort((short) 32767);
        row.setPrimInt(Integer.MIN_VALUE);
        row.setWrapInt(Integer.MAX_VALUE);
        row.setPrimLong(9007199254740991L);     // 2^53 - 1 (the largest integer exactly representable as a double)
        row.setWrapLong(-9007199254740991L);

        // Floating-point types (exactly representable values)
        row.setPrimDouble(1234.5);
        row.setWrapDouble(-0.25);
        row.setPrimFloat(1.5F);
        row.setWrapFloat(-2.25F);

        // Character
        row.setPrimChar('A');
        row.setWrapChar('z');

        // Boolean
        row.setPrimBool(true);
        row.setWrapBool(false);

        // Large numbers (exported as text cells so precision is preserved even beyond 2^53)
        row.setBigInt(new BigInteger("9999999999999999"));
        row.setBigDec(new BigDecimal("12345.6789"));

        // Date/time
        row.setJavaDate(Date.from(BASE_DATE_TIME.atZone(zone).toInstant()));
        row.setLocalDate(LocalDate.of(2023, 6, 15));
        row.setLocalTime(LocalTime.of(10, 30, 45));
        row.setLocalDateTime(BASE_DATE_TIME);
        row.setZonedDateTime(BASE_DATE_TIME.atZone(zone));
        row.setOffsetTime(LocalTime.of(10, 30, 45).atOffset(OffsetTime.now(zone).getOffset()));
        row.setOffsetDateTime(BASE_DATE_TIME.atZone(zone).toOffsetDateTime());
        row.setDuration(Duration.ofHours(1).plusMinutes(2).plusSeconds(3));
        row.setPeriod(Period.of(1, 2, 3));

        // UUID (upper-case hexadecimal digits, so the round-trip also shows the canonical lower-case normalization)
        row.setUuid(UUID.fromString("123E4567-E89B-12D3-A456-426614174000"));

        // enum
        row.setGrade(Grade.A);
        row.setCategory(Category.FOOD);

        // Custom objects
        row.setPoint(new Point(3, 7));
        row.setMoney(new Money("USD", 1050));

        // Collection
        row.setStringList(Arrays.asList("Apple", "Banana", "Cherry"));
        row.setIntList(Arrays.asList(10, 20, 30));
        row.setGradeList(Arrays.asList(Grade.A, Grade.B, Grade.F));

        return row;
    }

    /**
     * Verifies that the values created by sampleAllTypesRow() are all preserved after a round-trip (export -> import).
     */
    public static void assertSampleAllTypesRow(final AllTypesRow row) {

        final AllTypesRow expected = sampleAllTypesRow();

        assertThat(row).as("row was not created").isNotNull();
        assertThat(row.getRowIndex()).as("row index was not injected").isNotNull();

        assertThat(row.getText()).isEqualTo(expected.getText());
        assertThat(row.getLeadingZero()).as("leading-zero string was not preserved").isEqualTo("007");

        assertThat(row.getPrimByte()).isEqualTo(expected.getPrimByte());
        assertThat(row.getWrapByte()).isEqualTo(expected.getWrapByte());
        assertThat(row.getPrimShort()).isEqualTo(expected.getPrimShort());
        assertThat(row.getWrapShort()).isEqualTo(expected.getWrapShort());
        assertThat(row.getPrimInt()).isEqualTo(expected.getPrimInt());
        assertThat(row.getWrapInt()).isEqualTo(expected.getWrapInt());
        assertThat(row.getPrimLong()).isEqualTo(expected.getPrimLong());
        assertThat(row.getWrapLong()).isEqualTo(expected.getWrapLong());

        assertThat(row.getPrimDouble()).isEqualTo(expected.getPrimDouble());
        assertThat(row.getWrapDouble()).isEqualTo(expected.getWrapDouble());
        assertThat(row.getPrimFloat()).isEqualTo(expected.getPrimFloat());
        assertThat(row.getWrapFloat()).isEqualTo(expected.getWrapFloat());

        assertThat(row.getPrimChar()).isEqualTo(expected.getPrimChar());
        assertThat(row.getWrapChar()).isEqualTo(expected.getWrapChar());

        assertThat(row.isPrimBool()).isEqualTo(expected.isPrimBool());
        assertThat(row.getWrapBool()).isEqualTo(expected.getWrapBool());

        assertThat(row.getBigInt()).as("BigInteger precision was not preserved").isEqualTo(expected.getBigInt());
        assertThat(row.getBigDec()).as("BigDecimal value was not preserved").isEqualByComparingTo(expected.getBigDec());

        assertThat(row.getJavaDate()).isEqualTo(expected.getJavaDate());
        assertThat(row.getLocalDate()).isEqualTo(expected.getLocalDate());
        assertThat(row.getLocalTime()).isEqualTo(expected.getLocalTime());
        assertThat(row.getLocalDateTime()).isEqualTo(expected.getLocalDateTime());
        assertThat(row.getZonedDateTime()).isEqualTo(expected.getZonedDateTime());
        assertThat(row.getOffsetTime()).isEqualTo(expected.getOffsetTime());
        assertThat(row.getOffsetDateTime()).isEqualTo(expected.getOffsetDateTime());
        assertThat(row.getDuration()).isEqualTo(expected.getDuration());
        assertThat(row.getPeriod()).isEqualTo(expected.getPeriod());

        // The value is case-insensitive, so the round-trip through the canonical lower-case text preserves it.
        assertThat(row.getUuid()).as("UUID value was not preserved").isEqualTo(expected.getUuid());

        assertThat(row.getGrade()).isEqualTo(Grade.A);
        assertThat(row.getCategory()).isEqualTo(Category.FOOD);

        assertThat(row.getPoint()).as("custom object (Point) was not preserved").isNotNull();
        assertThat(row.getPoint().getX()).isEqualTo(3);
        assertThat(row.getPoint().getY()).isEqualTo(7);

        assertThat(row.getMoney()).as("custom object (Money) was not preserved").isNotNull();
        assertThat(row.getMoney().getCurrency()).isEqualTo("USD");
        assertThat(row.getMoney().getAmount()).isEqualTo(1050L);

        assertThat(row.getStringList()).containsExactly("Apple", "Banana", "Cherry");
        assertThat(row.getIntList()).containsExactly(10, 20, 30);
        assertThat(row.getGradeList()).containsExactly(Grade.A, Grade.B, Grade.F);
    }

    /**
     * Creates a minimal baseline row with every field filled with a valid, safe value.
     * <p>
     * This is the base for type-specific fine-grained scenario tests where only a particular field is varied for verification.
     * (Every field is filled with a real value so that an abnormal value like NUL in a primitive char does not break the export.)
     */
    public static AllTypesRow baseAllTypesRow() {

        final ZoneId zone = ZoneId.systemDefault();

        final AllTypesRow row = new AllTypesRow();

        row.setText("base");
        row.setLeadingZero("1");

        row.setPrimByte((byte) 1);
        row.setWrapByte((byte) 2);
        row.setPrimShort((short) 3);
        row.setWrapShort((short) 4);
        row.setPrimInt(5);
        row.setWrapInt(6);
        row.setPrimLong(7L);
        row.setWrapLong(8L);

        row.setPrimDouble(1.5);
        row.setWrapDouble(2.5);
        row.setPrimFloat(3.5F);
        row.setWrapFloat(4.5F);

        row.setPrimChar('A');
        row.setWrapChar('B');

        row.setPrimBool(true);
        row.setWrapBool(false);

        row.setBigInt(BigInteger.ONE);
        row.setBigDec(new BigDecimal("1.5"));

        row.setJavaDate(Date.from(BASE_DATE_TIME.atZone(zone).toInstant()));
        row.setLocalDate(LocalDate.of(2023, 6, 15));
        row.setLocalTime(LocalTime.of(10, 30, 45));
        row.setLocalDateTime(BASE_DATE_TIME);
        row.setZonedDateTime(BASE_DATE_TIME.atZone(zone));
        row.setOffsetTime(LocalTime.of(10, 30, 45).atOffset(OffsetTime.now(zone).getOffset()));
        row.setOffsetDateTime(BASE_DATE_TIME.atZone(zone).toOffsetDateTime());
        row.setDuration(Duration.ofSeconds(1));
        row.setPeriod(Period.ofDays(1));

        row.setUuid(UUID.fromString("00000000-0000-0000-0000-000000000001"));

        row.setGrade(Grade.A);
        row.setCategory(Category.FOOD);

        row.setPoint(new Point(1, 1));
        row.setMoney(new Money("USD", 1));

        row.setStringList(Arrays.asList("x"));
        row.setIntList(Arrays.asList(1));
        row.setGradeList(Arrays.asList(Grade.A));

        return row;
    }

    /**
     * Creates a simple Employee instance.
     */
    public static Employee employee(final String name,
                                    final int age,
                                    final String salary,
                                    final boolean active,
                                    final LocalDate hireDate,
                                    final Grade grade,
                                    final String department) {

        final Employee employee = new Employee();
        employee.setName(name);
        employee.setAge(age);
        employee.setSalary(new BigDecimal(salary));
        employee.setActive(active);
        employee.setHireDate(hireDate);
        employee.setGrade(grade);
        employee.setDepartment(department);

        return employee;
    }

    /**
     * Creates a simple Department instance.
     */
    public static Department department(final String code,
                                        final String name,
                                        final int headcount) {

        final Department department = new Department();
        department.setCode(code);
        department.setDepartmentName(name);
        department.setHeadcount(headcount);

        return department;
    }

}
