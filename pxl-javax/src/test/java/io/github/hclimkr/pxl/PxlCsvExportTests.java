package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.builder.PxlCsvExportBuilder;
import io.github.hclimkr.pxl.exception.*;
import io.github.hclimkr.pxl.option.PxlExportSheetOption;
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import io.github.hclimkr.pxl.option.PxlImportSheetOption;
import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;
import io.github.hclimkr.pxl.styler.data.PxlDataVerticalCenterTextStyler;
import io.github.hclimkr.pxl.tcdata.*;
import io.github.hclimkr.pxl.type.PxlExcelEngine;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInfo;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;

import java.io.*;
import java.math.BigDecimal;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.time.LocalDate;
import java.util.*;

import static io.github.hclimkr.pxl.tcdata.TestExports.emit;
import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * CSV export tests.
 * <p>
 * Covers {@link Pxl#exportCsv()}: the records written, the charset/delimiter/BOM cascade, the row and column
 * coordinates, the settings CSV ignores, and the guards the terminal raises.
 * <p>
 * What an export produces has to be the same on either terminal, so a test whose subject is the exported result is
 * swept across {@link ExportDest} with {@link TestExports#emit}. CSV has no {@code toWorkbook()}, so the sweep is
 * narrowed to {@code FILE} and {@code STREAM}. What stays a plain {@code @Test} is the opposite kind: a test whose
 * subject <em>is</em> one destination's mechanics - a caller's stream left open, a file that never appears, a
 * temporary spill file removed after the file could not be opened.
 * <p>
 * Encoding-sensitive expectations are asserted on the raw bytes rather than on a decoded string: decoding first
 * would absorb a byte order mark into U+FEFF and hide whether one was written at all.
 */
public class PxlCsvExportTests {

    private static Pxl pxl;

    @BeforeAll
    public static void setUpBeforeClass() {
        pxl = new Pxl();
    }

    // Captures the current test method name to match the export file name.
    private TestInfo testInfo;

    @BeforeEach
    public void bindTestInfo(final TestInfo testInfo) {
        this.testInfo = testInfo;
    }

    private File csvFile() {
        return TestPaths.exportFile(testInfo, ".csv");
    }

    private static List<Employee> twoEmployees() {
        final Employee alice = new Employee();
        alice.setName("Alice");
        alice.setAge(30);
        alice.setSalary(new BigDecimal("50000.50"));
        alice.setActive(Boolean.TRUE);
        alice.setHireDate(LocalDate.of(2020, 1, 15));
        alice.setGrade(Grade.A);
        alice.setDepartment("Engineering");

        final Employee bob = new Employee();
        bob.setName("Bob");
        bob.setAge(42);
        bob.setSalary(new BigDecimal("72000.00"));
        bob.setActive(Boolean.FALSE);
        bob.setHireDate(LocalDate.of(2018, 7, 1));
        bob.setGrade(Grade.B);
        bob.setDepartment("Sales");

        return Arrays.asList(alice, bob);
    }

    // Reads the file as raw bytes, which is the only way to tell a missing byte order mark from a decoded one.
    private static byte[] bytesOf(final File file) throws Exception {
        return Files.readAllBytes(file.toPath());
    }

    private static List<String> linesOf(final byte[] bytes, final Charset charset) {
        return Arrays.asList(new String(bytes, charset).split("\r\n", -1));
    }

    private static List<String> linesOf(final byte[] bytes) {
        return linesOf(bytes, StandardCharsets.UTF_8);
    }

    private static List<String> linesOf(final File file) throws Exception {
        return linesOf(bytesOf(file));
    }

    // Parses the output rather than splitting it: a field carrying the delimiter is quoted, so counting fields by
    // hand would read one value as several.
    private static List<CSVRecord> recordsOf(final byte[] bytes) throws Exception {
        try (Reader reader = new InputStreamReader(new ByteArrayInputStream(bytes), StandardCharsets.UTF_8);
             CSVParser parser = CSVParser.parse(reader, CSVFormat.EXCEL)) {
            return parser.getRecords();
        }
    }

    private static PxlExportWorkbookOption sheetOptionOf(final PxlExportSheetOption sheetOption) throws Exception {
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder().build();
        option.addExportSheetOption(sheetOption);
        return option;
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_sheetForm_writesHeaderAndDataRecords(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees"), dest, testInfo);

        final List<String> lines = linesOf(bytes);
        // The header is always written: CSV export has no switch to turn it off.
        assertThat(lines.get(0)).isEqualTo("Name,Age,Salary,Active,HireDate,Grade,Department");
        assertThat(lines.get(1)).isEqualTo("Alice,30,50000.50,true,2020-01-15,A,Engineering");
        assertThat(lines.get(2)).isEqualTo("Bob,42,72000.00,false,2018-07-01,B,Sales");
        // Trailing record separator only; no second header and no extra record.
        assertThat(lines.get(3)).isEmpty();
        assertThat(lines).hasSize(4);
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_uuidColumns_writeCanonicalText(final ExportDest dest) throws Exception {
        final String uuidText = "123e4567-e89b-12d3-a456-426614174000";
        final String otherUuidText = "00112233-4455-6677-8899-aabbccddeeff";

        final UuidRow row = new UuidRow();
        row.setId(UUID.fromString(uuidText));
        row.setIds(Arrays.asList(UUID.fromString(otherUuidText), UUID.fromString(uuidText)));

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(UuidRow.class, Arrays.asList(row), "Uuids"), dest, testInfo);

        final List<String> lines = linesOf(bytes);
        assertThat(lines.get(0)).isEqualTo("Id,Ids,Exact,Masked,Unique");
        // The collection separator is not the delimiter, so the joined field needs no quoting; the unset columns are
        // written with their exportNullString (the empty string by default).
        assertThat(lines.get(1)).isEqualTo(uuidText + "," + otherUuidText + ";" + uuidText + ",,,");
    }

    // Not swept: only the stream destination is handed a stream it does not own.
    @Test
    public void exportCsv_toStream_flushesEverythingAndLeavesStreamOpen() throws Exception {
        final boolean[] closed = {false};
        final ByteArrayOutputStream buffer = new ByteArrayOutputStream();
        final OutputStream watched = new OutputStream() {

            @Override
            public void write(final int b) {
                buffer.write(b);
            }

            @Override
            public void write(final byte[] b, final int off, final int len) {
                buffer.write(b, off, len);
            }

            @Override
            public void close() {
                closed[0] = true;
            }
        };

        pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .toStream(watched);

        // Missing a flush anywhere in the chain would truncate the output silently.
        assertThat(new String(buffer.toByteArray(), StandardCharsets.UTF_8))
                .startsWith("Name,Age,")
                .contains("Alice,30,")
                .contains("Bob,42,");
        assertThat(closed[0]).isFalse();
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_delimiterOption_writesTabSeparatedFields(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvDelimiter('\t')
                        .build()), dest, testInfo);

        assertThat(linesOf(bytes).get(0)).isEqualTo("Name\tAge\tSalary\tActive\tHireDate\tGrade\tDepartment");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_charsetOption_encodesWithGivenCharset(final ExportDest dest) throws Exception {
        final CharsetRow row = new CharsetRow();
        row.setCity("Seoul");

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(CharsetRow.class, Collections.singletonList(row), "Cities")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvCharset("EUC-KR")
                        .build()), dest, testInfo);

        assertThat(linesOf(bytes, Charset.forName("EUC-KR")).get(1)).isEqualTo("Seoul");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_sheetOptionCharset_overridesWorkbookOption(final ExportDest dest) throws Exception {
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportCsvCharset("UTF-8")
                .exportCsvDelimiter(',')
                .build();
        option.addExportSheetOption(PxlExportSheetOption.builder()
                .exportCsvCharset("UTF-16LE")
                .exportCsvDelimiter(';')
                .build());

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(option), dest, testInfo);

        // The sheet level wins over the workbook level, exactly as it does on import.
        assertThat(linesOf(bytes, StandardCharsets.UTF_16LE).get(0))
                .isEqualTo("Name;Age;Salary;Active;HireDate;Grade;Department");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_sheetOptionLeavingCsvValuesUnset_inheritsTheWorkbookLevel(final ExportDest dest) throws Exception {
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportCsvCharset("UTF-16LE")
                .exportCsvDelimiter(';')
                .build();
        // A sheet option present but silent on the CSV values must not shadow the workbook's with its own blanks.
        option.addExportSheetOption(PxlExportSheetOption.builder()
                .exportIfEmpty(true)
                .build());

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(option), dest, testInfo);

        assertThat(linesOf(bytes, StandardCharsets.UTF_16LE).get(0))
                .isEqualTo("Name;Age;Salary;Active;HireDate;Grade;Department");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_sheetOptionBom_overridesWorkbookOption(final ExportDest dest) throws Exception {
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportCsvBom(true)
                .build();
        // A byte order mark belongs to the file, so the sheet level settles it just as it settles the charset.
        option.addExportSheetOption(PxlExportSheetOption.builder()
                .exportCsvBom(false)
                .build());

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(option), dest, testInfo);

        assertThat(bytes[0]).isEqualTo((byte) 'N');
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_sheetOptionLeavingBomUnset_inheritsTheWorkbookLevel(final ExportDest dest) throws Exception {
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportCsvBom(true)
                .build();
        option.addExportSheetOption(PxlExportSheetOption.builder()
                .exportCsvCharset("UTF-8")
                .build());

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(option), dest, testInfo);

        assertThat(Arrays.copyOf(bytes, 3)).containsExactly((byte) 0xEF, (byte) 0xBB, (byte) 0xBF);
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_withoutBomOption_writesNoByteOrderMark(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees"), dest, testInfo);

        // The default is off, so the file must open on the header rather than on a mark.
        assertThat(bytes[0]).isEqualTo((byte) 'N');
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_bomUtf8_writesByteOrderMark(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvBom(true)
                        .build()), dest, testInfo);

        assertThat(Arrays.copyOf(bytes, 3)).containsExactly((byte) 0xEF, (byte) 0xBB, (byte) 0xBF);
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_bomCharsetAlias_stillWritesByteOrderMark(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvCharset("utf8")   // an alias, not the canonical name
                        .exportCsvBom(true)
                        .build()), dest, testInfo);

        // Comparing charset names instead of Charset instances would drop the mark here, and since a dropped mark
        // is normal behaviour elsewhere the bug would look like the feature.
        assertThat(Arrays.copyOf(bytes, 3)).containsExactly((byte) 0xEF, (byte) 0xBB, (byte) 0xBF);
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_bomUtf16Le_writesByteOrderMark(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvCharset("UTF-16LE")
                        .exportCsvBom(true)
                        .build()), dest, testInfo);

        assertThat(Arrays.copyOf(bytes, 2)).containsExactly((byte) 0xFF, (byte) 0xFE);
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_bomUtf16Be_writesByteOrderMark(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvCharset("UTF-16BE")
                        .exportCsvBom(true)
                        .build()), dest, testInfo);

        assertThat(Arrays.copyOf(bytes, 2)).containsExactly((byte) 0xFE, (byte) 0xFF);
        assertThat(linesOf(bytes, StandardCharsets.UTF_16BE).get(0)).endsWith("Department");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_bomUtf16_writesTheMarkOnlyOnce(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvCharset("UTF-16")   // the encoder writes a mark of its own
                        .exportCsvBom(true)
                        .build()), dest, testInfo);

        assertThat(Arrays.copyOf(bytes, 2)).containsExactly((byte) 0xFE, (byte) 0xFF);
        // A mark written on top of the encoder's would show up as U+FEFF ahead of the first header.
        assertThat(new String(bytes, "UTF-16")).startsWith("Name");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_bomNonUnicodeCharset_skipsMarkAndKeepsFirstHeaderReadable(final ExportDest dest) throws Exception {
        final CharsetRow row = new CharsetRow();
        row.setCity("Seoul");

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(CharsetRow.class, Collections.singletonList(row), "Cities")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvCharset("EUC-KR")
                        .exportCsvBom(true)
                        .build()), dest, testInfo);

        // U+FEFF has no EUC-KR encoding and would become '?', corrupting the first header cell, so it is dropped
        // silently - no exception, no log.
        assertThat(bytes[0]).isNotEqualTo((byte) '?');
        assertThat(linesOf(bytes, Charset.forName("EUC-KR")).get(0)).isEqualTo("City");

        // The point of dropping it: the file PXL wrote is one PXL can read back.
        final List<CharsetRow> loaded = pxl.importCsv()
                .sheet(CharsetRow.class)
                .override(PxlImportWorkbookOption.builder()
                        .importCsvCharset("EUC-KR")
                        .build())
                .fromStream("Cities", new ByteArrayInputStream(bytes));
        assertThat(loaded).extracting(CharsetRow::getCity).containsExactly("Seoul");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_nullValue_writesExportNullString(final ExportDest dest) throws Exception {
        final NullStringRow row = new NullStringRow();
        row.setValue(null);
        row.setLabel(null);

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(NullStringRow.class, Collections.singletonList(row), "Nulls"), dest, testInfo);

        // The configured null string for the first column, the default empty string for the second - and both
        // still occupy a field, so the record keeps the header's field count.
        assertThat(linesOf(bytes).get(1)).isEqualTo("N/A,");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_exportOrder_writesColumnsInDeclaredOrder(final ExportDest dest) throws Exception {
        final OrderedRow row = new OrderedRow();
        row.setX("x");
        row.setY("y");
        row.setZ("z");

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(OrderedRow.class, Collections.singletonList(row), "Ordered"), dest, testInfo);

        final List<String> lines = linesOf(bytes);
        assertThat(lines.get(0)).isEqualTo("Y,Z,X");
        assertThat(lines.get(1)).isEqualTo("y,z,x");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_disabledColumn_isOmittedFromEveryRecord(final ExportDest dest) throws Exception {
        final ColumnToggleRow row = new ColumnToggleRow();
        row.setAlways("a");
        row.setExportOff("b");
        row.setImportOff("c");

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(ColumnToggleRow.class, Collections.singletonList(row), "Toggle"), dest, testInfo);

        final List<String> lines = linesOf(bytes);
        assertThat(lines.get(0)).isEqualTo("Always,ImportOff");
        assertThat(lines.get(1)).isEqualTo("a,c");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_maskingColumn_writesMaskedString(final ExportDest dest) throws Exception {
        final MaskingRow row = new MaskingRow();
        row.setSecret("abc123");

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(MaskingRow.class, Collections.singletonList(row), "Masked"), dest, testInfo);

        assertThat(linesOf(bytes).get(1)).isEqualTo("abc***");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_everySupportedType_writesAValueForEveryColumn(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(AllTypesRow.class, Collections.singletonList(Fixtures.sampleAllTypesRow()), "AllTypes")
                .override(Fixtures.noValidationOption()), dest, testInfo);

        final List<CSVRecord> records = recordsOf(bytes);
        // Every codec is reached with no cell to write into - the primitives among them are a path the Collection
        // codec never covers - and each still occupies its own field.
        assertThat(records).hasSize(2);
        assertThat(records.get(1).size()).isEqualTo(records.get(0).size());
        assertThat(records.get(1).get(0)).isEqualTo("Hello, PXL");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_rowAndColumnCoordinates_writeLeadingEmptyFieldRecords(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(sheetOptionOf(PxlExportSheetOption.builder()
                        .exportHeaderRowIndex(3)
                        .exportFirstDataColumnIndex(2)
                        .build())), dest, testInfo);

        final List<String> lines = linesOf(bytes);
        // Two records stand in for the rows above the header. They must not be blank lines: PXL's own import
        // ignores those, which would pull the header up and break the round trip below.
        assertThat(lines.get(0)).isEqualTo("\"\",,,,,,,");
        assertThat(lines.get(1)).isEqualTo("\"\",,,,,,,");
        assertThat(lines.get(2)).isEqualTo("\"\",Name,Age,Salary,Active,HireDate,Grade,Department");
        assertThat(lines.get(3)).isEqualTo("\"\",Alice,30,50000.50,true,2020-01-15,A,Engineering");

        // The written coordinates are the ones import reads back.
        final PxlImportWorkbookOption importOption = PxlImportWorkbookOption.builder().build();
        importOption.addImportSheetOption(PxlImportSheetOption.builder()
                .importHeaderRowIndex(3)
                .importFirstDataColumnIndex(2)
                .build());

        final List<Employee> loaded = pxl.importCsv()
                .sheet(Employee.class)
                .override(importOption)
                .fromStream("Employees", new ByteArrayInputStream(bytes));
        assertThat(loaded).extracting(Employee::getName).containsExactly("Alice", "Bob");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_lastDataColumnIndexTooSmall_throwsLikeExcel(final ExportDest dest) {
        // The bound is a guard rather than a truncation: leaving an enabled column unmapped is reported instead of
        // silently dropping it, and that decision is made in the shared column metadata, so CSV inherits it.
        assertThrows(PxlDataException.class, () -> emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(sheetOptionOf(PxlExportSheetOption.builder()
                        .exportLastDataColumnIndex(3)   // fewer than the seven columns Employee binds
                        .build())), dest, testInfo));
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_i18nResourceBundle_translatesColumnNames(final ExportDest dest) throws Exception {
        final I18nRow row = new I18nRow();
        row.setFullName("Alice");
        row.setRole("Admin");

        // The base bundle is asked for without the JVM's default-locale fallback, the way PXL resolves one itself:
        // messages_en does not exist, and letting the fallback run would pick up whatever locale the machine has.
        final ResourceBundle bundle = ResourceBundle.getBundle("messages", Locale.ENGLISH,
                ResourceBundle.Control.getNoFallbackControl(ResourceBundle.Control.FORMAT_PROPERTIES));

        // The sheet form reads no workbook annotation, so the bundle comes through the option - the one i18n path
        // a CSV export can actually take today.
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(I18nRow.class, Collections.singletonList(row), "Staff")
                .override(PxlExportWorkbookOption.builder()
                        .exportResourceBundle(bundle)
                        .exportDataValidation(false)
                        .build()), dest, testInfo);

        // The column names are keys: staff.column.fullName -> "Full Name", staff.column.role -> "Role".
        assertThat(recordsOf(bytes).get(0).toList()).containsExactlyInAnyOrder("Full Name", "Role");
        // Data values are not translated, so the row is written as given.
        assertThat(recordsOf(bytes).get(1).toList()).contains("Alice", "Admin");
    }

    // Not swept: only the file destination can fail to open.
    @Test
    public void exportCsv_toFileDestinationUnopenable_throwsAndBuilderStaysUsable() throws Exception {
        final PxlCsvExportBuilder builder = pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees");

        final File unopenable = new File(TestPaths.EXPORT_DIR + "/no-such-directory/out.csv");
        // Opening the file fails with an IOException, which the terminal normalizes.
        assertThrows(PxlSystemException.class, () -> builder.toFile(unopenable));
        assertThat(unopenable).doesNotExist();

        // The buffer prepared for the failed run was released in the terminal's finally, so the same builder still
        // renders afresh. Releasing outside that finally would leave the whole output on the heap instead.
        final File csvFile = csvFile();
        builder.toFile(csvFile);
        assertThat(linesOf(csvFile).get(1)).startsWith("Alice,");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_lastDataRowIndex_truncatesDataRecords(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(sheetOptionOf(PxlExportSheetOption.builder()
                        .exportLastDataRowIndex(2)   // header on row 1, so a single data row fits
                        .build())), dest, testInfo);

        final List<String> lines = linesOf(bytes);
        assertThat(lines.get(1)).startsWith("Alice,");
        assertThat(lines.get(2)).isEmpty();
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_trimColumn_writesTheTrimmedValue(final ExportDest dest) throws Exception {
        final TrimRow row = new TrimRow();
        row.setPadded("  padded  ");

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(TrimRow.class, Collections.singletonList(row), "Trim"), dest, testInfo);

        assertThat(linesOf(bytes).get(1)).isEqualTo("padded");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_inheritedColumns_areWrittenLikeExcel(final ExportDest dest) throws Exception {
        final DerivedRow row = new DerivedRow();
        row.setId(7);
        row.setBaseName("base");
        row.setExtra("extra");

        // The column metadata walks the class hierarchy, and CSV shares that layer with Excel.
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(DerivedRow.class, Collections.singletonList(row), "Derived"), dest, testInfo);

        assertThat(recordsOf(bytes).get(0).toList()).containsExactlyInAnyOrder("Id", "BaseName", "Extra");
        assertThat(recordsOf(bytes).get(1).toList()).containsExactlyInAnyOrder("7", "base", "extra");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_dateTimeWithoutPattern_writesTheCodecStringNotASerial(final ExportDest dest) throws Exception {
        final DateTimeNumericRow row = new DateTimeNumericRow();
        row.setLocalDate(LocalDate.of(2023, 6, 15));

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(DateTimeNumericRow.class, Collections.singletonList(row), "Dates")
                .override(Fixtures.noValidationOption()), dest, testInfo);

        // Excel writes a patternless date as a numeric serial carrying a display format; CSV has neither cell types
        // nor styles, so what lands in the field is the codec's own string - which is why a date survives the CSV
        // round trip more plainly than the Excel one.
        final int dateIndex = recordsOf(bytes).get(0).toList().indexOf("LocalDate");
        assertThat(recordsOf(bytes).get(1).get(dateIndex)).isEqualTo("2023-06-15");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_stringAsFormula_writesTheTextItself(final ExportDest dest) throws Exception {
        final FormulaRow row = new FormulaRow();
        row.setLabel("Sum");
        row.setFormula("=1+2");

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(FormulaRow.class, Collections.singletonList(row), "Formula"), dest, testInfo);

        // Nothing is evaluated and nothing is dropped: the leading '=' is part of the value, so the field carries
        // the expression verbatim. "Ignored" would be the wrong word for this attribute.
        final int formulaIndex = recordsOf(bytes).get(0).toList().indexOf("Formula");
        assertThat(recordsOf(bytes).get(1).get(formulaIndex)).isEqualTo("=1+2");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_stringAsPicture_writesTheImageLocation(final ExportDest dest) throws Exception {
        final PictureRow row = new PictureRow();
        row.setPhoto("images/photo.png");

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(PictureRow.class, Collections.singletonList(row), "Pictures"), dest, testInfo);

        // No picture is embedded, so the location the value held is what reaches the field. Worth pinning: a
        // reader expecting "ignored" would not expect the path to be disclosed.
        final int photoIndex = recordsOf(bytes).get(0).toList().indexOf("Photo");
        assertThat(recordsOf(bytes).get(1).get(photoIndex)).isEqualTo("images/photo.png");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_dropdownAndWidthSettings_areIgnoredWithoutFailing(final ExportDest dest) throws Exception {
        final EnumOptionItemsRow row = new EnumOptionItemsRow();

        // exportOptionItems / exportEnumDropDownListStyle / exportColumnWidth have no CSV counterpart, and asking
        // for them must not fail the export.
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(EnumOptionItemsRow.class, Collections.singletonList(row), "Dropdown")
                .override(Fixtures.noValidationOption()), dest, testInfo);

        assertThat(recordsOf(bytes)).hasSize(2);
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_excelOnlySettings_areIgnoredWithoutFailing(final ExportDest dest) throws Exception {
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportExcelEngine(PxlExcelEngine.SXSSF)
                .exportSXSSFRowAccessWindowSize(50)
                .exportWorkbookDataCellStyler(PxlDataVerticalCenterTextStyler.class)
                .build();
        option.addExportSheetOption(PxlExportSheetOption.builder()
                .exportColumnFilter(true)
                .exportRowHeightInPoints(30f)
                .exportGroupingFieldName("department")
                .build());

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(option), dest, testInfo);

        final List<String> lines = linesOf(bytes);
        assertThat(lines.get(0)).isEqualTo("Name,Age,Salary,Active,HireDate,Grade,Department");
        // Grouping splits sheets on the Excel path; CSV has one file, so the rows stay in the order given.
        assertThat(lines.get(1)).startsWith("Alice,");
        assertThat(lines.get(2)).startsWith("Bob,");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_secondSheet_failsAtTheTerminalNotAtSheet(final ExportDest dest) throws Exception {
        // Configuring a second sheet is accepted, just as it is for Excel...
        final PxlCsvExportBuilder builder = pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .sheet(Employee.class, twoEmployees(), "More");

        // ...and it is the terminal, which writes one destination, that refuses it - either terminal alike.
        assertThrows(PxlArgumentException.class, () -> emit(builder, dest, testInfo));
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_noSheet_throws(final ExportDest dest) {
        assertThrows(PxlArgumentException.class, () -> emit(pxl.exportCsv(), dest, testInfo));
    }

    // Not swept: these are rejected at the config step, before any terminal is in play.
    @Test
    public void exportCsv_invalidSheetArguments_areRejectedAtTheConfigStep() {
        // Required arguments are checked where they are given, matching the Excel builder rather than deferring to
        // the terminal the way the single-sheet rule does.
        assertThrows(PxlNullPointerException.class, () -> pxl.exportCsv()
                .sheet(null, twoEmployees(), "Employees"));
        assertThrows(PxlNullPointerException.class, () -> pxl.exportCsv()
                .sheet(Employee.class, null, "Employees"));
        assertThrows(PxlArgumentException.class, () -> pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), " "));
    }

    // Not swept: what is being pinned is that no plaintext file is left on disk.
    @Test
    public void exportCsv_exportPassword_throwsRatherThanWritingPlaintext() throws Exception {
        final File csvFile = csvFile();
        Files.deleteIfExists(csvFile.toPath());

        assertThrows(PxlArgumentException.class, () -> pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportPassword("secret")
                        .build())
                .toFile(csvFile));

        // Rejected before the destination was opened, so no plaintext file was left behind.
        assertThat(csvFile).doesNotExist();
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_invalidCharset_throws(final ExportDest dest) {
        assertThrows(PxlArgumentException.class, () -> emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvCharset("NoSuchCharset-1")
                        .build()), dest, testInfo));
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_invalidDelimiter_throws(final ExportDest dest) {
        assertThrows(PxlArgumentException.class, () -> emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvDelimiter('\n')
                        .build()), dest, testInfo));
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_delimiterEqualToQuoteChar_throws(final ExportDest dest) {
        // Commons-CSV judges a delimiter by whether it collides with the dialect's quote, escape or comment
        // character, so what counts as usable follows the dialect the printer is built from rather than the
        // character on its own.
        assertThrows(PxlArgumentException.class, () -> emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvDelimiter('"')
                        .build()), dest, testInfo));
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_valuesCarryingCsvSyntax_areQuotedAndSurviveTheRoundTrip(final ExportDest dest) throws Exception {
        final Employee awkward = new Employee();
        awkward.setName("Doe, John \"JD\"\nsecond line");
        awkward.setAge(1);
        awkward.setDepartment("R&D");

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, Collections.singletonList(awkward), "Employees"), dest, testInfo);

        // The delimiter, the quote and the newline all have to survive being written and read back.
        final List<CSVRecord> records = recordsOf(bytes);
        assertThat(records.get(1).get(0)).isEqualTo("Doe, John \"JD\"\nsecond line");

        final List<Employee> loaded = pxl.importCsv()
                .sheet(Employee.class)
                .fromStream("Employees", new ByteArrayInputStream(bytes));
        assertThat(loaded).extracting(Employee::getName).containsExactly("Doe, John \"JD\"\nsecond line");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_emptyRowsWithExportIfEmptyOn_writesHeaderOnlyAndReadsBackEmpty(final ExportDest dest) throws Exception {
        // The default keeps the sheet, which for CSV means a header-only file: a blank form that can be filled in.
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, new ArrayList<Employee>(), "Employees"), dest, testInfo);

        final List<String> lines = linesOf(bytes);
        assertThat(lines.get(0)).isEqualTo("Name,Age,Salary,Active,HireDate,Grade,Department");
        assertThat(lines.get(1)).isEmpty();

        assertThat(pxl.importCsv()
                .sheet(Employee.class)
                .fromStream("Employees", new ByteArrayInputStream(bytes)))
                .isEmpty();
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_nullRowInTheCollection_throwsLikeExcel(final ExportDest dest) {
        final List<Employee> rows = new ArrayList<>(twoEmployees());
        rows.add(null);

        // Skipping it would drop a row without saying so, so CSV fails exactly where Excel does.
        assertThrows(PxlDataException.class, () -> emit(pxl.exportCsv()
                .sheet(Employee.class, rows, "Employees"), dest, testInfo));
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_beanValidationViolation_throws(final ExportDest dest) {
        final ValidatedRow row = new ValidatedRow();
        row.setName(null);      // @NotBlank violation
        row.setAge(20);

        // Bean validation is on by default and runs before a single record is rendered.
        assertThrows(PxlValidationException.class, () -> emit(pxl.exportCsv()
                .sheet(ValidatedRow.class, Collections.singletonList(row), "V"), dest, testInfo));
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_beanValidationDisabled_writesTheViolatingRow(final ExportDest dest) throws Exception {
        final ValidatedRow row = new ValidatedRow();
        row.setName(null);
        row.setAge(20);

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(ValidatedRow.class, Collections.singletonList(row), "V")
                .override(Fixtures.noValidationOption()), dest, testInfo);

        assertThat(linesOf(bytes)).hasSizeGreaterThan(1);
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_sheetEnabledFlagInTheSheetForm_isIgnoredLikeExcel(final ExportDest dest) throws Exception {
        // The sheet form takes the row class as an argument, so there is no sheet field to switch off and the sheet
        // meta hardcodes exportEnabled. The option therefore cannot suppress the sheet here - on the Excel sheet
        // form it cannot either. It becomes meaningful only once a workbook form exists.
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(sheetOptionOf(PxlExportSheetOption.builder()
                        .exportEnabled(false)
                        .build())), dest, testInfo);

        assertThat(linesOf(bytes).get(0)).isEqualTo("Name,Age,Salary,Active,HireDate,Grade,Department");
        assertThat(linesOf(bytes).get(1)).startsWith("Alice,");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_rowClassBindingNoColumn_throws(final ExportDest dest) {
        // Matches the Excel path, which reports the same state as "nothing to write".
        assertThrows(PxlDataException.class, () -> emit(pxl.exportCsv()
                .sheet(NoColumnRow.class, Collections.singletonList(new NoColumnRow()), "Empty"), dest, testInfo));
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_emptyRowsWithExportIfEmptyOff_throws(final ExportDest dest) {
        assertThrows(PxlDataException.class, () -> emit(pxl.exportCsv()
                .sheet(Employee.class, new ArrayList<Employee>(), "Employees")
                .override(sheetOptionOf(PxlExportSheetOption.builder()
                        .exportIfEmpty(false)
                        .build())), dest, testInfo));
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_emptyRowsWithExportIfEmptyOn_writesHeaderOnly(final ExportDest dest) throws Exception {
        // The default keeps an empty collection exportable, and what lands on the destination is the header record
        // alone - a file the import side can still read back as zero rows.
        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, new ArrayList<Employee>(), "Employees"), dest, testInfo);

        assertThat(recordsOf(bytes)).hasSize(1);
        assertThat(linesOf(bytes).get(0)).isEqualTo("Name,Age,Salary,Active,HireDate,Grade,Department");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_rowCountExceeded_throws(final ExportDest dest) {
        // The data bound counts from the first data row, so a high enough origin exceeds the CSV row cap without
        // actually holding a hundred thousand rows.
        assertThrows(PxlDataException.class, () -> emit(pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees")
                .override(sheetOptionOf(PxlExportSheetOption.builder()
                        .exportFirstDataRowIndex(PxlConstants.EXPORT_MAX_NUMBER_OF_CSV_ROWS + 1)
                        .build())), dest, testInfo));
    }

    // Not swept: what is being pinned is that no partial file is left on disk.
    @Test
    public void exportCsv_codecFailureOnLaterRow_leavesNoFileAndBuilderStaysUsable() throws Exception {
        final File csvFile = csvFile();
        Files.deleteIfExists(csvFile.toPath());

        // The first row encodes cleanly and the second does not: a non-finite double is rejected by the codec, at
        // encode time rather than while the metadata is built. The failure therefore lands after the header and one
        // data record have been produced, which is the case that tells a buffered writer from a streaming one.
        final AllTypesRow encodable = Fixtures.sampleAllTypesRow();
        final AllTypesRow failing = Fixtures.sampleAllTypesRow();
        failing.setWrapDouble(Double.NaN);

        assertThrows(PxlCellCodecException.class, () -> pxl.exportCsv()
                .sheet(AllTypesRow.class, Arrays.asList(encodable, failing), "AllTypes")
                .override(Fixtures.noValidationOption())
                .toFile(csvFile));

        // The records are rendered before the destination is opened, so a failure halfway through leaves neither a
        // truncated file nor an empty one. Streaming the records straight out would break this first.
        assertThat(csvFile).doesNotExist();

        // The same builder still works afterwards.
        final PxlCsvExportBuilder builder = pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees");
        builder.toFile(csvFile);
        assertThat(linesOf(csvFile).get(1)).startsWith("Alice,");
    }

    // Not swept: the subject is one builder run twice, so the destination is the fixture.
    @Test
    public void exportCsv_rerun_buildsFreshOutputEachTime() throws Exception {
        final PxlCsvExportBuilder builder = pxl.exportCsv()
                .sheet(Employee.class, twoEmployees(), "Employees");

        final ByteArrayOutputStream first = new ByteArrayOutputStream();
        final ByteArrayOutputStream second = new ByteArrayOutputStream();
        builder.toStream(first);
        builder.toStream(second);

        // A cached result would make the second run either empty or a duplicate of the first buffer.
        assertThat(second.toByteArray()).isEqualTo(first.toByteArray());
        assertThat(new String(second.toByteArray(), StandardCharsets.UTF_8)).startsWith("Name,Age,");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_everyRecord_carriesTheSameFieldCountAsTheHeader(final ExportDest dest) throws Exception {
        final Employee blank = new Employee();

        final List<Employee> rows = new ArrayList<>(twoEmployees());
        rows.add(blank);

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(Employee.class, rows, "Employees"), dest, testInfo);

        final List<CSVRecord> records = recordsOf(bytes);
        final int headerFields = records.get(0).size();
        for (int index = 1; index < records.size(); index++) {
            // A mapped column keeps its field even when the codec answers null, or every later column of that
            // record would shift one place left.
            assertThat(records.get(index).size()).isEqualTo(headerFields);
        }
    }

    // ------------------------------------------------------------------
    // Output too large to hold in memory spills to a temporary file
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_outputOverTheMemoryThreshold_spillsToDiskAndWritesEveryRecord(final ExportDest dest) throws Exception {
        final List<LargeTextRow> rows = largeTextRows();

        final int temporariesBefore = countSpillFiles();

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(LargeTextRow.class, rows, "Large"), dest, testInfo);

        // Past the threshold the render carries on into a temporary file, so this exercises the spilled path and
        // not the in-memory one. Asserting on the size rather than on the file itself keeps the test honest if the
        // threshold moves: the output either passed it or the test is no longer testing a spill.
        assertThat(bytes.length).isGreaterThan(PxlConstants.EXPORT_MEMORY_THRESHOLD_OF_CSV);

        final List<CSVRecord> records = recordsOf(bytes);
        assertThat(records).hasSize(rows.size() + 1);
        assertThat(records.get(0).toList()).containsExactly("Text", "Value");
        // The hand-over from memory to file happens part-way through, so a record after it is the one that would
        // show the buffered part being dropped or written twice.
        assertThat(records.get(records.size() - 1).get(0)).isEqualTo(rows.get(0).getText());

        // The temporary file belongs to the run, not to the builder.
        assertThat(countSpillFiles()).isEqualTo(temporariesBefore);
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportCsv_spilledOutput_readsBackAsItWasWritten(final ExportDest dest) throws Exception {
        final List<LargeTextRow> rows = largeTextRows();

        final byte[] bytes = emit(pxl.exportCsv()
                .sheet(LargeTextRow.class, rows, "Large"), dest, testInfo);

        final List<LargeTextRow> readBack = pxl.importCsv()
                .sheet(LargeTextRow.class)
                .fromStream("Large", new ByteArrayInputStream(bytes));

        assertThat(readBack).hasSize(rows.size());
        assertThat(readBack.get(0).getText()).isEqualTo(rows.get(0).getText());
        assertThat(readBack.get(readBack.size() - 1).getText()).isEqualTo(rows.get(0).getText());
    }

    // Not swept: only the file destination can fail to open, and the spill file it leaves behind is the point.
    @Test
    public void exportCsv_spilledRenderWithDestinationUnopenable_leavesNoTemporaryFile() throws Exception {
        final List<LargeTextRow> rows = largeTextRows();
        final File unopenable = new File(TestPaths.EXPORT_DIR + "/no-such-directory/out.csv");

        final int temporariesBefore = countSpillFiles();

        assertThrows(PxlSystemException.class, () -> pxl.exportCsv()
                .sheet(LargeTextRow.class, rows, "Large")
                .toFile(unopenable));

        assertThat(unopenable).doesNotExist();
        // The render had already spilled by the time the destination failed to open, so the terminal's finally is
        // what has to remove the temporary file.
        assertThat(countSpillFiles()).isEqualTo(temporariesBefore);
    }

    // Not swept: what is being pinned is that neither the destination file nor the spill file is left behind.
    @Test
    public void exportCsv_codecFailureAfterSpilling_leavesNoFileAndNoTemporaryFile() throws Exception {
        final File csvFile = csvFile();
        Files.deleteIfExists(csvFile.toPath());

        final List<LargeTextRow> rows = new ArrayList<>(largeTextRows());
        final LargeTextRow failing = new LargeTextRow();
        failing.setText("last");
        // A non-finite double is rejected at encode time, so the failure lands once the render is well past the
        // threshold and a temporary file is already open.
        failing.setValue(Double.NaN);
        rows.add(failing);

        final int temporariesBefore = countSpillFiles();

        assertThrows(PxlCellCodecException.class, () -> pxl.exportCsv()
                .sheet(LargeTextRow.class, rows, "Large")
                .toFile(csvFile));

        // Rendering still finishes before the destination is opened, spill or no spill.
        assertThat(csvFile).doesNotExist();
        // Nothing closed the sink on this path, which is why cleanup() has to be able to reach a sink that was
        // abandoned mid-render.
        assertThat(countSpillFiles()).isEqualTo(temporariesBefore);
    }

    // The name a CSV export gives the temporary file it spills into.
    private static final String SPILL_FILE_PREFIX = "pxl-csv-export-";

    private static int countSpillFiles() {
        final String[] found = new File(System.getProperty("java.io.tmpdir"))
                .list((dir, name) -> name.startsWith(SPILL_FILE_PREFIX));

        return found == null ? 0 : found.length;
    }

    // Enough rows to carry the output past the memory threshold, with one shared String behind every row so the
    // rows themselves stay small while the output does not.
    private static List<LargeTextRow> largeTextRows() {
        final int textWidth = 10_000;
        final int rowCount = PxlConstants.EXPORT_MEMORY_THRESHOLD_OF_CSV / textWidth + 100;

        final char[] text = new char[textWidth];
        Arrays.fill(text, 'x');
        final String shared = new String(text);

        final List<LargeTextRow> rows = new ArrayList<>(rowCount);
        for (int index = 0; index < rowCount; index++) {
            final LargeTextRow row = new LargeTextRow();
            row.setText(shared);
            row.setValue(1.0);
            rows.add(row);
        }

        return rows;
    }

}
