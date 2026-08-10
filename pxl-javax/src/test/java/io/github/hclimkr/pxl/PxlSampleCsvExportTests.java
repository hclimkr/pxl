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

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * CSV sample (template) export tests.
 * <p>
 * Covers {@link Pxl#exportSampleCsv()}, which writes a header record plus a single record filled from each
 * column's {@code exportSample} value. The template's practical worth is that it can be filled in and read back,
 * so the round trip is asserted here too.
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

    private static List<String> linesOf(final File file) throws Exception {
        final String text = new String(Files.readAllBytes(file.toPath()), StandardCharsets.UTF_8);
        return Arrays.asList(text.split("\r\n", -1));
    }

    @Test
    public void exportSampleCsv_sheetForm_writesHeaderAndOneSampleRecord() throws Exception {
        final File csvFile = csvFile();

        pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .toFile(csvFile);

        final List<String> lines = linesOf(csvFile);
        assertThat(lines.get(0)).isEqualTo("Name,Age,Salary,Active,HireDate,Grade,Department");
        assertThat(lines.get(1)).isEqualTo("Alice,30,50000.50,true,2020-01-15,A,Engineering");
        // Exactly one sample record, then the trailing separator.
        assertThat(lines.get(2)).isEmpty();
        assertThat(lines).hasSize(3);
    }

    @Test
    public void exportSampleCsv_disabledSampleColumn_isOmitted() throws Exception {
        final File csvFile = csvFile();

        pxl.exportSampleCsv()
                .sheet(SampleColumnRow.class, "Samples")
                .toFile(csvFile);

        final List<String> lines = linesOf(csvFile);
        assertThat(lines.get(0)).isEqualTo("Keep");
        assertThat(lines.get(1)).isEqualTo("K");
    }

    @Test
    public void exportSampleCsv_blankSample_stillOccupiesItsField() throws Exception {
        final File csvFile = csvFile();

        // Only the first column declares a sample; the second must still take a field so the columns stay aligned.
        pxl.exportSampleCsv()
                .sheet(NullStringRow.class, "Nulls")
                .toFile(csvFile);

        final List<String> lines = linesOf(csvFile);
        assertThat(lines.get(0).split(",", -1)).hasSize(2);
        assertThat(lines.get(1).split(",", -1)).hasSize(2);
    }

    @Test
    public void exportSampleCsv_rowClassBindingNoColumn_throws() {
        // No column metadata at all, which is the state the Excel sample path reports as "nothing to write".
        assertThrows(PxlDataException.class, () -> pxl.exportSampleCsv()
                .sheet(NoColumnRow.class, "Empty")
                .toStream(new ByteArrayOutputStream()));
    }

    @Test
    public void exportSampleCsv_everyColumnOptedOutOfSample_writesEmptyRecordsLikeExcel() throws Exception {
        final ByteArrayOutputStream buffer = new ByteArrayOutputStream();

        // Opting out of the sample does not remove a column from the metadata, it only leaves it unmapped, so this
        // is not the "nothing to write" case above. The Excel sample path answers it with a sheet holding two
        // cell-less rows, and CSV matches that with two empty records rather than failing.
        pxl.exportSampleCsv()
                .sheet(NoSampleColumnRow.class, "Empty")
                .toStream(buffer);

        assertThat(new String(buffer.toByteArray(), StandardCharsets.UTF_8)).isEqualTo("\r\n\r\n");
    }

    @Test
    public void exportSampleCsv_generatedTemplate_isReadBackByImportCsv() throws Exception {
        final File csvFile = csvFile();

        pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .toFile(csvFile);

        final List<Employee> loaded = pxl.importCsv()
                .sheet(Employee.class)
                .fromFile(csvFile);

        // The declared sample values bind straight back, which is what makes the template usable as a form.
        assertThat(loaded).hasSize(1);
        assertThat(loaded.get(0).getName()).isEqualTo("Alice");
        assertThat(loaded.get(0).getAge()).isEqualTo(30);
        assertThat(loaded.get(0).getHireDate()).isEqualTo(LocalDate.of(2020, 1, 15));
        assertThat(loaded.get(0).getGrade()).isEqualTo(Grade.A);
    }

    @Test
    public void exportSampleCsv_delimiterOption_writesTabSeparatedFields() throws Exception {
        final File csvFile = csvFile();

        pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvDelimiter('\t')
                        .build())
                .toFile(csvFile);

        assertThat(linesOf(csvFile).get(0)).isEqualTo("Name\tAge\tSalary\tActive\tHireDate\tGrade\tDepartment");
    }

    @Test
    public void exportSampleCsv_rowAndColumnCoordinates_writeLeadingEmptyFieldRecords() throws Exception {
        final File csvFile = csvFile();
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder().build();
        option.addExportSheetOption(PxlExportSheetOption.builder()
                .exportHeaderRowIndex(2)
                .exportFirstDataColumnIndex(2)
                .build());

        pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .override(option)
                .toFile(csvFile);

        final List<String> lines = linesOf(csvFile);
        // One record stands in for the row above the header, and it is an empty-field record rather than a blank
        // line, which is what keeps the template readable back.
        assertThat(lines.get(0)).isEqualTo("\"\",,,,,,,");
        assertThat(lines.get(1)).isEqualTo("\"\",Name,Age,Salary,Active,HireDate,Grade,Department");
        assertThat(lines.get(2)).isEqualTo("\"\",Alice,30,50000.50,true,2020-01-15,A,Engineering");
        assertThat(lines.get(3)).isEmpty();
    }

    @Test
    public void exportSampleCsv_charsetOption_encodesWithGivenCharset() throws Exception {
        final File csvFile = csvFile();

        pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvCharset("UTF-16LE")
                        .build())
                .toFile(csvFile);

        final String text = new String(Files.readAllBytes(csvFile.toPath()), StandardCharsets.UTF_16LE);
        assertThat(text).startsWith("Name,Age,");
    }

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

    @Test
    public void exportSampleCsv_bomUtf8_writesByteOrderMark() throws Exception {
        final File csvFile = csvFile();

        pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvBom(true)
                        .build())
                .toFile(csvFile);

        final byte[] bytes = Files.readAllBytes(csvFile.toPath());
        assertThat(Arrays.copyOf(bytes, 3)).containsExactly((byte) 0xEF, (byte) 0xBB, (byte) 0xBF);
    }

    // The guards live on the shared CSV base, so they are asserted on both builders: a regression that only
    // reaches one of them would otherwise pass here.

    @Test
    public void exportSampleCsv_secondSheet_failsAtTheTerminalNotAtSheet() throws Exception {
        final PxlSampleCsvExportBuilder builder = pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .sheet(Employee.class, "More");

        assertThrows(PxlArgumentException.class, () -> builder.toStream(new ByteArrayOutputStream()));
    }

    @Test
    public void exportSampleCsv_noSheet_throws() {
        assertThrows(PxlArgumentException.class, () -> pxl.exportSampleCsv().toStream(new ByteArrayOutputStream()));
    }

    @Test
    public void exportSampleCsv_invalidSheetArguments_areRejectedAtTheConfigStep() {
        assertThrows(PxlNullPointerException.class, () -> pxl.exportSampleCsv().sheet(null, "Employees"));
        assertThrows(PxlArgumentException.class, () -> pxl.exportSampleCsv().sheet(Employee.class, " "));
    }

    @Test
    public void exportSampleCsv_everyColumnType_getsASampleValue() throws Exception {
        final File csvFile = csvFile();

        pxl.exportSampleCsv()
                .sheet(AllTypesRow.class, "AllTypes")
                .toFile(csvFile);

        // Every codec has to render its declared sample with no cell to write into, and each still occupies a field.
        try (Reader reader = new InputStreamReader(new ByteArrayInputStream(Files.readAllBytes(csvFile.toPath())), StandardCharsets.UTF_8);
             CSVParser parser = CSVParser.parse(reader, CSVFormat.EXCEL)) {
            final List<CSVRecord> records = parser.getRecords();
            assertThat(records).hasSize(2);
            assertThat(records.get(1).size()).isEqualTo(records.get(0).size());
            assertThat(records.get(1).toList()).doesNotContainNull();
        }
    }

    @Test
    public void exportSampleCsv_enumSampleNotAnEnumConstant_throws() {
        // Reverse-parsing an enum sample is a shared codec rule, so the sample CSV fails exactly where the sample
        // Excel does.
        assertThrows(PxlCellCodecException.class, () -> pxl.exportSampleCsv()
                .sheet(BadEnumSampleRow.class, "Bad")
                .toStream(new ByteArrayOutputStream()));
    }

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

    @Test
    public void exportSampleCsv_invalidCharset_throws() {
        assertThrows(PxlArgumentException.class, () -> pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvCharset("NoSuchCharset-1")
                        .build())
                .toStream(new ByteArrayOutputStream()));
    }

    @Test
    public void exportSampleCsv_invalidDelimiter_throws() {
        assertThrows(PxlArgumentException.class, () -> pxl.exportSampleCsv()
                .sheet(Employee.class, "Employees")
                .override(PxlExportWorkbookOption.builder()
                        .exportCsvDelimiter('"')   // the same character as the quote
                        .build())
                .toStream(new ByteArrayOutputStream()));
    }

    // ------------------------------------------------------------------
    // The sample record is always written, whatever exportLastDataRowIndex declares
    // ------------------------------------------------------------------

    @Test
    public void exportSampleCsv_lastDataRowIndexBeforeFirstDataRow_stillWritesSampleRecord() throws Exception {
        // A sample carries exactly one data record whatever the declared bound says. With the header on 0-based
        // row 0 the sample lands on row 1, so a declared bound of 1 (1-based) points at the header row -- ahead of
        // the record actually written. The counterpart of the Excel test, which asserts the same on the ranges
        // the bound feeds there.
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportSheetOptions(Arrays.asList(PxlExportSheetOption.builder()
                        .exportLastDataRowIndex(1)
                        .build()))
                .build();

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        pxl.exportSampleCsv()
                .sheet(SampleDropdownRow.class, "Sample")
                .override(option)
                .toStream(outputStream);

        final List<String> lines = Arrays.asList(new String(outputStream.toByteArray(), StandardCharsets.UTF_8).split("\r\n", -1));

        assertThat(lines.get(0)).isEqualTo("Name,Choice");
        assertThat(lines.get(1)).as("the sample record must be written even so").isEqualTo("Alice,Red");
    }

}
