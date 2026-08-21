package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.builder.PxlSampleCsvExportBuilder;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.exception.PxlDataException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.option.PxlExportSheetOption;
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import io.github.hclimkr.pxl.tcdata.*;
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
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;

import static io.github.hclimkr.pxl.tcdata.TestExports.emit;
import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * CSV sample (template) export tests.
 * <p>
 * Covers {@link Pxl#exportSampleCsv()}, which writes a header record plus a single record filled from each
 * column's {@code exportSample} value. The template's practical worth is that it can be filled in and read back,
 * so the round trip is asserted here too.
 * <p>
 * A template has to come out the same on either terminal, so a test whose subject is the written template is swept
 * across {@link ExportDest} with {@link TestExports#emit} - narrowed to {@code FILE} and {@code STREAM}, since CSV
 * has no {@code toWorkbook()}. What stays a plain {@code @Test} is a test whose subject <em>is</em> one
 * destination's mechanics, or one that never reaches a terminal at all.
 */
public class PxlSampleCsvExportTests {

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

    private static List<String> linesOf(final byte[] bytes, final Charset charset) {
        return Arrays.asList(new String(bytes, charset).split("\r\n", -1));
    }

    private static List<String> linesOf(final byte[] bytes) {
        return linesOf(bytes, StandardCharsets.UTF_8);
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_sheetForm_writesHeaderAndOneSampleRecord(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees"), dest, testInfo);

        final List<String> lines = linesOf(bytes);
        assertThat(lines.get(0)).isEqualTo("Name,Age,Salary,Active,HireDate,Grade,Department");
        assertThat(lines.get(1)).isEqualTo("Alice,30,50000.50,true,2020-01-15,A,Engineering");
        // Exactly one sample record, then the trailing separator.
        assertThat(lines.get(2)).isEmpty();
        assertThat(lines).hasSize(3);
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_disabledSampleColumn_isOmitted(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportSampleCsv()
                .sheet(SampleColumnRow.class, "Samples"), dest, testInfo);

        final List<String> lines = linesOf(bytes);
        assertThat(lines.get(0)).isEqualTo("Keep");
        assertThat(lines.get(1)).isEqualTo("K");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_blankSample_stillOccupiesItsField(final ExportDest dest) throws Exception {
        // Only the first column declares a sample; the second must still take a field so the columns stay aligned.
        final byte[] bytes = emit(pxl.exportSampleCsv()
                .sheet(NullStringRow.class, "Nulls"), dest, testInfo);

        final List<String> lines = linesOf(bytes);
        assertThat(lines.get(0).split(",", -1)).hasSize(2);
        assertThat(lines.get(1).split(",", -1)).hasSize(2);
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_rowClassBindingNoColumn_throws(final ExportDest dest) {
        // No column metadata at all, which is the state the Excel sample path reports as "nothing to write".
        assertThrows(PxlDataException.class, () -> emit(pxl.exportSampleCsv()
                .sheet(NoColumnRow.class, "Empty"), dest, testInfo));
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_everyColumnOptedOutOfSample_writesEmptyRecordsLikeExcel(final ExportDest dest) throws Exception {
        // Opting out of the sample does not remove a column from the metadata, it only leaves it unmapped, so this
        // is not the "nothing to write" case above. The Excel sample path answers it with a sheet holding two
        // cell-less rows, and CSV matches that with two empty records rather than failing.
        final byte[] bytes = emit(pxl.exportSampleCsv()
                .sheet(NoSampleColumnRow.class, "Empty"), dest, testInfo);

        assertThat(new String(bytes, StandardCharsets.UTF_8)).isEqualTo("\r\n\r\n");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_generatedTemplate_isReadBackByImportCsv(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees"), dest, testInfo);

        final List<Employee> loaded = pxl.importCsv()
                .sheet(Employee.class)
                .fromStream("Employees", new ByteArrayInputStream(bytes));

        // The declared sample values bind straight back, which is what makes the template usable as a form.
        assertThat(loaded).hasSize(1);
        assertThat(loaded.get(0).getName()).isEqualTo("Alice");
        assertThat(loaded.get(0).getAge()).isEqualTo(30);
        assertThat(loaded.get(0).getHireDate()).isEqualTo(LocalDate.of(2020, 1, 15));
        assertThat(loaded.get(0).getGrade()).isEqualTo(Grade.A);
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_delimiterOption_writesTabSeparatedFields(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvDelimiter('\t')
                        .build()), dest, testInfo);

        assertThat(linesOf(bytes).get(0)).isEqualTo("Name\tAge\tSalary\tActive\tHireDate\tGrade\tDepartment");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_rowAndColumnCoordinates_writeLeadingEmptyFieldRecords(final ExportDest dest) throws Exception {
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder().build();
        option.addExportSheetOption(PxlExportSheetOption.builder()
                .exportHeaderRowIndex(2)
                .exportFirstDataColumnIndex(2)
                .build());

        final byte[] bytes = emit(pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .override(option), dest, testInfo);

        final List<String> lines = linesOf(bytes);
        // One record stands in for the row above the header, and it is an empty-field record rather than a blank
        // line, which is what keeps the template readable back.
        assertThat(lines.get(0)).isEqualTo("\"\",,,,,,,");
        assertThat(lines.get(1)).isEqualTo("\"\",Name,Age,Salary,Active,HireDate,Grade,Department");
        assertThat(lines.get(2)).isEqualTo("\"\",Alice,30,50000.50,true,2020-01-15,A,Engineering");
        assertThat(lines.get(3)).isEmpty();
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_charsetOption_encodesWithGivenCharset(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvCharset("UTF-16LE")
                        .build()), dest, testInfo);

        assertThat(new String(bytes, StandardCharsets.UTF_16LE)).startsWith("Name,Age,");
    }

    // Not swept: the subject is one builder run twice, so the destination is the fixture.
    @Test
    public void exportSampleCsv_rerun_buildsFreshOutputEachTime() throws Exception {
        final PxlSampleCsvExportBuilder builder = pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees");

        final ByteArrayOutputStream first = new ByteArrayOutputStream();
        final ByteArrayOutputStream second = new ByteArrayOutputStream();
        builder.toStream(first);
        builder.toStream(second);

        // A cached buffer would leave the second run empty, and one that is never released would grow it.
        assertThat(second.toByteArray()).isEqualTo(first.toByteArray());
        assertThat(new String(second.toByteArray(), StandardCharsets.UTF_8)).startsWith("Name,Age,");
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_bomUtf8_writesByteOrderMark(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvBom(true)
                        .build()), dest, testInfo);

        assertThat(Arrays.copyOf(bytes, 3)).containsExactly((byte) 0xEF, (byte) 0xBB, (byte) 0xBF);
    }

    // The guards live on the shared CSV base, so they are asserted on both builders: a regression that only
    // reaches one of them would otherwise pass here.

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_secondSheet_failsAtTheTerminalNotAtSheet(final ExportDest dest) throws Exception {
        final PxlSampleCsvExportBuilder builder = pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .sheet(Employee.class, "More");

        assertThrows(PxlArgumentException.class, () -> emit(builder, dest, testInfo));
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_noSheet_throws(final ExportDest dest) {
        assertThrows(PxlArgumentException.class, () -> emit(pxl.exportSampleCsv(), dest, testInfo));
    }

    // Not swept: these are rejected at the config step, before any terminal is in play.
    @Test
    public void exportSampleCsv_invalidSheetArguments_areRejectedAtTheConfigStep() {
        assertThrows(PxlNullPointerException.class, () -> pxl.exportSampleCsv().sheet(null, "Employees"));
        assertThrows(PxlArgumentException.class, () -> pxl.exportSampleCsv().sheet(Employee.class, " "));
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_everyColumnType_getsASampleValue(final ExportDest dest) throws Exception {
        final byte[] bytes = emit(pxl.exportSampleCsv()
                .sheet(AllTypesRow.class, "AllTypes"), dest, testInfo);

        // Every codec has to render its declared sample with no cell to write into, and each still occupies a field.
        try (Reader reader = new InputStreamReader(new ByteArrayInputStream(bytes), StandardCharsets.UTF_8);
             CSVParser parser = CSVParser.parse(reader, CSVFormat.EXCEL)) {
            final List<CSVRecord> records = parser.getRecords();
            assertThat(records).hasSize(2);
            assertThat(records.get(1).size()).isEqualTo(records.get(0).size());
            assertThat(records.get(1).toList()).doesNotContainNull();
        }
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_enumSampleNotAnEnumConstant_throws(final ExportDest dest) {
        // Reverse-parsing an enum sample is a shared codec rule, so the sample CSV fails exactly where the sample
        // Excel does.
        assertThrows(PxlCellCodecException.class, () -> emit(pxl.exportSampleCsv()
                .sheet(BadEnumSampleRow.class, "Bad"), dest, testInfo));
    }

    // Not swept: what is being pinned is that no plaintext file is left on disk.
    @Test
    public void exportSampleCsv_exportPassword_throwsRatherThanWritingPlaintext() throws Exception {
        final File csvFile = csvFile();
        Files.deleteIfExists(csvFile.toPath());

        assertThrows(PxlArgumentException.class, () -> pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportPassword("secret")
                        .build())
                .toFile(csvFile));

        assertThat(csvFile).doesNotExist();
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_invalidCharset_throws(final ExportDest dest) {
        assertThrows(PxlArgumentException.class, () -> emit(pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvCharset("NoSuchCharset-1")
                        .build()), dest, testInfo));
    }

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_invalidDelimiter_throws(final ExportDest dest) {
        assertThrows(PxlArgumentException.class, () -> emit(pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvDelimiter('"')   // the same character as the quote
                        .build()), dest, testInfo));
    }

    // ------------------------------------------------------------------
    // The sample record is always written, whatever exportLastDataRowIndex declares
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})
    public void exportSampleCsv_lastDataRowIndexBeforeFirstDataRow_stillWritesSampleRecord(final ExportDest dest) throws Exception {
        // A sample carries exactly one data record whatever the declared bound says. With the header on 0-based
        // row 0 the sample lands on row 1, so a declared bound of 1 (1-based) points at the header row -- ahead of
        // the record actually written. The counterpart of the Excel test, which asserts the same on the ranges
        // the bound feeds there.
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportSheetOptions(Arrays.asList(PxlExportSheetOption.builder()
                        .exportLastDataRowIndex(1)
                        .build()))
                .build();

        final byte[] bytes = emit(pxl.exportSampleCsv()
                .sheet(SampleDropdownRow.class, "Sample")
                .override(option), dest, testInfo);

        final List<String> lines = linesOf(bytes);

        assertThat(lines.get(0)).isEqualTo("Name,Choice");
        assertThat(lines.get(1)).as("the sample record must be written even so").isEqualTo("Alice,Red");
    }

}
