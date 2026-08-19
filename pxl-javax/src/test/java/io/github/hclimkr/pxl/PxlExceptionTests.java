package io.github.hclimkr.pxl;

import com.github.pjfanning.xlsx.StreamingReader;
import io.github.hclimkr.pxl.exception.*;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.option.*;
import io.github.hclimkr.pxl.tcdata.*;
import io.github.hclimkr.pxl.util.PxlCellUtils;
import io.github.hclimkr.pxl.util.PxlMiscUtils;
import io.github.hclimkr.pxl.util.PxlSheetUtils;
import io.github.hclimkr.pxl.util.PxlWorkbookUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInfo;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.*;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.*;

/**
 * Exception scenario tests - file/validation/special double/duplicate sheet/password/stream-reader formula, plus type parsing failures (bool/bigdecimal/biginteger/localdate/duration/period), integer/byte out-of-range, ERROR cell, no default constructor, RowIndex type/unsupported type/grouping field typo, and which validation pass reports a violation when a @PxlSheet field carries @Valid.
 */
public class PxlExceptionTests {

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

    private static InputStream stream(final String content) {
        return new ByteArrayInputStream(content.getBytes(StandardCharsets.US_ASCII));
    }

    // Sheet bytes with 1 header row + 1 STRING data row
    private static byte[] sheetWithValue(final String header, final String value) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final Sheet sheet = workbook.createSheet("Data");
            sheet.createRow(0).createCell(0).setCellValue(header);
            sheet.createRow(1).createCell(0).setCellValue(value);
            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    // Imports the "Data" sheet with the given rowClass (for exception checks)
    private static void importTyped(final byte[] bytes, final Class<?> rowClass) throws Exception {
        pxl.importExcel()
                .sheet(rowClass, Arrays.asList("Data"))
                .fromStream(new ByteArrayInputStream(bytes));
    }

    private static void assertImportThrows(final String header, final String value) {
        assertThrows(PxlCellCodecException.class, () -> importTyped(sheetWithValue(header, value), TypedRow.class));
    }

    // ------------------------------------------------------------------
    // import: non-existent file
    // ------------------------------------------------------------------

    @Test
    public void importFile_nonExistent_throws(@TempDir final Path tempDir) {
        final File missing = tempDir.resolve("does-not-exist.xlsx").toFile();

        assertThrows(PxlIOException.class, () ->
                pxl.importExcel()
                        .sheet(Employee.class, Arrays.asList("Any"))
                        .fromFile(missing));
    }

    // ------------------------------------------------------------------
    // export: Bean Validation violation (required value missing)
    // ------------------------------------------------------------------

    @Test
    public void exportValidation_blankRequiredField_throws() {
        final ValidatedRow row = new ValidatedRow();
        row.setName(null);      // @NotBlank violation
        row.setAge(20);

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        // Data validation is enabled by default (null option), so validation is performed on export.
        assertThrows(PxlValidationException.class, () ->
                pxl.exportExcel()
                        .sheet(ValidatedRow.class, Arrays.asList(row), "V")
                        .toStream(outputStream));
    }

    @Test
    public void exportValidation_nullRequiredField_throws() {
        final ValidatedRow row = new ValidatedRow();
        row.setName("Alice");
        row.setAge(null);       // @NotNull violation

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlValidationException.class, () ->
                pxl.exportExcel()
                        .sheet(ValidatedRow.class, Arrays.asList(row), "V")
                        .toStream(outputStream));
    }

    // ------------------------------------------------------------------
    // @Valid on a @PxlSheet field: which of the two validation passes reports the violation
    // ------------------------------------------------------------------

    // Sheet bytes with a "Name","Age" header row + 1 data row, matching CountingRow.
    private static byte[] nameAgeSheet(final String sheetName, final String name, final String age) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final Sheet sheet = workbook.createSheet(sheetName);

            final Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Name");
            header.createCell(1).setCellValue("Age");

            final Row data = sheet.createRow(1);
            data.createCell(0).setCellValue(name);
            data.createCell(1).setCellValue(age);

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    // Sheet bytes with the "Name","Age" header row and no data row at all.
    private static byte[] headerOnlySheet(final String sheetName) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final Sheet sheet = workbook.createSheet(sheetName);

            final Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Name");
            header.createCell(1).setCellValue("Age");

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    private static CountingRow blankNameRow() {
        final CountingRow row = new CountingRow();
        row.setName("   ");     // @CountedNotBlank violation
        row.setAge(30);

        return row;
    }

    @Test
    public void exportWorkbook_cascadingSheetField_violatingRow_throwsWithSheetTag() {
        final CascadeWorkbook workbook = new CascadeWorkbook();
        workbook.setWorkbookName("Cascade");
        workbook.setRows(Arrays.asList(blankNameRow()));

        try {
            // The location tag is localized against the JVM default locale, so pin it to read the tag itself.
            Pxl.setMessageLocale(Locale.ENGLISH);

            // The workbook-object pass runs first on export and carries no location, so it used to report this one
            // untagged. With the cascade skipped it no longer sees the row, and the tagged pass reports it instead.
            final PxlValidationException ex = assertThrows(PxlValidationException.class, () ->
                    pxl.exportExcel()
                            .workbook(workbook)
                            .toStream(new ByteArrayOutputStream()));
            assertThat(ex.getMessage()).contains("sheet 'Rows'");
        } finally {
            Pxl.resetMessageLocale();
        }
    }

    @Test
    public void exportWorkbook_nonCascadingSheetField_violatingRow_throwsWithSheetTag() {
        final NoCascadeWorkbook workbook = new NoCascadeWorkbook();
        workbook.setWorkbookName("NoCascade");
        workbook.setRows(Arrays.asList(blankNameRow()));

        try {
            Pxl.setMessageLocale(Locale.ENGLISH);

            // Without the cascade the row collection pass is the one that catches it, and it does tag the sheet.
            final PxlValidationException ex = assertThrows(PxlValidationException.class, () ->
                    pxl.exportExcel()
                            .workbook(workbook)
                            .toStream(new ByteArrayOutputStream()));
            assertThat(ex.getMessage()).contains("sheet 'Rows'");
            assertThat(ex.getMessage()).doesNotContain("row ");
        } finally {
            Pxl.resetMessageLocale();
        }
    }

    @Test
    public void importWorkbook_cascadingSheetField_violatingRow_throwsWithRowTag() throws Exception {
        final byte[] bytes = nameAgeSheet("Rows", "   ", "30");

        try {
            Pxl.setMessageLocale(Locale.ENGLISH);

            // On import the per-row pass runs first, so the violation is reported with sheet and row. The header is
            // row 1, hence the data row is reported as row 2.
            final PxlValidationException ex = assertThrows(PxlValidationException.class, () ->
                    pxl.importExcel()
                            .workbook(CascadeWorkbook.class)
                            .fromStream(new ByteArrayInputStream(bytes)));
            assertThat(ex.getMessage()).contains("sheet 'Rows'");
            assertThat(ex.getMessage()).contains("row 2");
        } finally {
            Pxl.resetMessageLocale();
        }
    }

    @Test
    public void importWorkbook_nonCascadingSheetField_violatingRow_throwsWithRowTag() throws Exception {
        final byte[] bytes = nameAgeSheet("Rows", "   ", "30");

        try {
            Pxl.setMessageLocale(Locale.ENGLISH);

            // @Valid changes nothing here: on import the per-row pass is first either way.
            final PxlValidationException ex = assertThrows(PxlValidationException.class, () ->
                    pxl.importExcel()
                            .workbook(NoCascadeWorkbook.class)
                            .fromStream(new ByteArrayInputStream(bytes)));
            assertThat(ex.getMessage()).contains("sheet 'Rows'");
            assertThat(ex.getMessage()).contains("row 2");
        } finally {
            Pxl.resetMessageLocale();
        }
    }

    // ------------------------------------------------------------------
    // Skipping the sheet cascade must not take any other constraint away
    // ------------------------------------------------------------------

    @Test
    public void exportWorkbook_sheetCollectionConstraint_stillValidated() {
        final ConstrainedCascadeWorkbook workbook = new ConstrainedCascadeWorkbook();
        workbook.setWorkbookName("Constrained");
        workbook.setRows(Collections.emptyList());      // @NotEmpty violation on the sheet field itself

        // The skipped cascade is only the descent into the elements. A constraint on the collection is a constraint
        // of the workbook object, and the workbook-object pass still checks it.
        final PxlValidationException ex = assertThrows(PxlValidationException.class, () ->
                pxl.exportExcel()
                        .workbook(workbook)
                        .toStream(new ByteArrayOutputStream()));
        assertThat(ex.getMessage()).contains("'Rows' must not be empty.");
    }

    @Test
    public void exportWorkbook_exportDisabledSheet_violatingRow_doesNotThrow() throws Exception {
        final CountingRow exported = new CountingRow();
        exported.setName("Alice");
        exported.setAge(30);

        final DisabledCascadeWorkbook workbook = new DisabledCascadeWorkbook();
        workbook.setWorkbookName("Disabled");
        workbook.setRows(Arrays.asList(exported));
        workbook.setSkippedRows(Arrays.asList(blankNameRow()));      // exportEnabled=false, so the binder skips it

        // A sheet that is not written is not validated either, @Valid on the field notwithstanding - so a violating
        // row there cannot fail an export that never includes it.
        pxl.exportExcel()
                .workbook(workbook)
                .toStream(new ByteArrayOutputStream());
    }

    @Test
    public void exportWorkbook_ownFieldConstraint_stillValidated() {
        final CountingRow row = new CountingRow();
        row.setName("Alice");
        row.setAge(30);

        final ConstrainedCascadeWorkbook workbook = new ConstrainedCascadeWorkbook();
        workbook.setWorkbookName("   ");                // @NotBlank violation on the workbook's own field
        workbook.setRows(Arrays.asList(row));

        // Nothing about the sheet rule touches a plain constraint on the workbook object itself.
        final PxlValidationException ex = assertThrows(PxlValidationException.class, () ->
                pxl.exportExcel()
                        .workbook(workbook)
                        .toStream(new ByteArrayOutputStream()));
        assertThat(ex.getMessage()).contains("'WorkbookName' must not be blank.");
    }

    @Test
    public void exportWorkbook_nonSheetFieldCascade_stillValidated() {
        final CountingRow row = new CountingRow();
        row.setName("Alice");
        row.setAge(30);

        final CascadeMeta meta = new CascadeMeta();
        meta.setLabel("   ");                           // violation behind a @Valid on a non-sheet field

        final ConstrainedCascadeWorkbook workbook = new ConstrainedCascadeWorkbook();
        workbook.setWorkbookName("Constrained");
        workbook.setRows(Arrays.asList(row));
        workbook.setMeta(meta);

        // Only @PxlSheet fields are exempt from cascading; every other @Valid still descends as the user declared.
        final PxlValidationException ex = assertThrows(PxlValidationException.class, () ->
                pxl.exportExcel()
                        .workbook(workbook)
                        .toStream(new ByteArrayOutputStream()));
        assertThat(ex.getMessage()).contains("'Label' must not be blank.");
    }

    @Test
    public void importWorkbook_ownFieldConstraint_stillValidated() throws Exception {
        // The export side has several of these; import needs its own, because there the workbook-object pass runs
        // last instead of first and could have been dropped without any export test noticing.
        final byte[] bytes = nameAgeSheet("Rows", "Alice", "30");

        // fromStream supplies no workbook name, so the field stays null and its @NotBlank is violated - by the pass
        // that runs after the rows are bound.
        final PxlValidationException ex = assertThrows(PxlValidationException.class, () ->
                pxl.importExcel()
                        .workbook(ConstrainedCascadeWorkbook.class)
                        .fromStream(new ByteArrayInputStream(bytes)));
        assertThat(ex.getMessage()).contains("'WorkbookName' must not be blank.");
    }

    @Test
    public void importWorkbook_sheetCollectionConstraint_stillValidated() throws Exception {
        // Header row only: the sheet binds to an empty collection, which its @NotEmpty rejects.
        final byte[] bytes = headerOnlySheet("Rows");

        final PxlValidationException ex = assertThrows(PxlValidationException.class, () ->
                pxl.importExcel()
                        .workbookName("Constrained")
                        .workbook(ConstrainedCascadeWorkbook.class)
                        .fromStream(new ByteArrayInputStream(bytes)));
        assertThat(ex.getMessage()).contains("'Rows' must not be empty.");
    }

    @Test
    public void importCsvWorkbook_ownFieldConstraint_stillValidated() throws Exception {
        // The CSV importer wires the same two passes, so the same guarantee has to hold there.
        final String csv = "Name,Age\nAlice,30\n";

        final PxlValidationException ex = assertThrows(PxlValidationException.class, () ->
                pxl.importCsv()
                        .workbook(ConstrainedCascadeWorkbook.class)
                        .fromStreams(Arrays.asList("Rows"), Arrays.asList(stream(csv))));
        assertThat(ex.getMessage()).contains("'WorkbookName' must not be blank.");
    }

    @Test
    public void exportWorkbook_nestedRowListCascade_stillValidated() {
        final CountingRow row = new CountingRow();
        row.setName("Alice");
        row.setAge(30);

        // The same row class, but reached below the root through a plain @Valid list rather than a sheet. The
        // binder never walks these, so the cascade has to reach them.
        final CascadeMeta meta = new CascadeMeta();
        meta.setLabel("Meta");
        meta.setNestedRows(Arrays.asList(blankNameRow()));

        final ConstrainedCascadeWorkbook workbook = new ConstrainedCascadeWorkbook();
        workbook.setWorkbookName("Constrained");
        workbook.setRows(Arrays.asList(row));
        workbook.setMeta(meta);

        final PxlValidationException ex = assertThrows(PxlValidationException.class, () ->
                pxl.exportExcel()
                        .workbook(workbook)
                        .toStream(new ByteArrayOutputStream()));
        assertThat(ex.getMessage()).contains("'Name' must not be blank.");
    }


    // ------------------------------------------------------------------
    // export: double NaN / Infinity fails fast
    // ------------------------------------------------------------------

    @Test
    public void exportDouble_nan_throws() {
        final AllTypesRow row = Fixtures.sampleAllTypesRow();
        row.setWrapDouble(Double.NaN);

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlCellCodecException.class, () ->
                pxl.exportExcel()
                        .sheet(AllTypesRow.class, Arrays.asList(row), "AllTypes")
                        .override(noValidationOption())
                        .toStream(outputStream));
    }

    @Test
    public void exportDouble_infinity_throws() {
        final AllTypesRow row = Fixtures.sampleAllTypesRow();
        row.setPrimDouble(Double.POSITIVE_INFINITY);

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlCellCodecException.class, () ->
                pxl.exportExcel()
                        .sheet(AllTypesRow.class, Arrays.asList(row), "AllTypes")
                        .override(noValidationOption())
                        .toStream(outputStream));
    }

    // ------------------------------------------------------------------
    // export: duplicate sheet names (explicit list form)
    // ------------------------------------------------------------------

    @Test
    public void exportWorkbook_duplicateSheetNames_throws() {
        final List<Employee> some = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, null, "Engineering"));

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        // Specifying the same sheet name ("Dup") twice -> exception
        assertThrows(PxlDataException.class, () ->
                pxl.exportExcel()
                        .sheet(Employee.class, some, "Dup")
                        .sheet(Employee.class, new ArrayList<Employee>(), "Dup")
                        .override(noValidationOption())
                        .toStream(outputStream));
    }

    @Test
    public void exportWorkbook_duplicateSheetNamesDifferentCase_throws() {
        // "Dup" and "DUP" are one sheet to a workbook, so the export is rejected up front rather than failing
        // later when POI refuses the second sheet - and the message names the offender.
        final List<Employee> some = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, null, "Engineering"));

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        final PxlDataException exception = assertThrows(PxlDataException.class, () ->
                pxl.exportExcel()
                        .sheet(Employee.class, some, "Dup")
                        .sheet(Employee.class, new ArrayList<Employee>(), "DUP")
                        .override(noValidationOption())
                        .toStream(outputStream));

        assertThat(exception).hasMessageContaining("DUP");
    }

    @Test
    public void exportSampleWorkbook_duplicateSheetNamesDifferentCase_throws() {
        // A sample export names its sheets the same way, so it is checked the same way.
        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        final PxlDataException exception = assertThrows(PxlDataException.class, () ->
                pxl.exportSampleExcel()
                        .sheet(Employee.class, "Dup")
                        .sheet(Employee.class, "DUP")
                        .toStream(outputStream));

        assertThat(exception).hasMessageContaining("DUP");
    }

    @Test
    public void exportWorkbookObject_duplicateSheetNamesDifferentCase_throws() {
        // The workbook form is checked on the names its @PxlSheet fields resolve to. Field order is not
        // guaranteed, so either of the two names may be reported as the duplicate.
        final List<Employee> some = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, null, "Engineering"));

        final DuplicateCaseSheetWorkbook workbook = new DuplicateCaseSheetWorkbook();
        workbook.setWorkbookName("W");
        workbook.setEmployees(some);
        workbook.setStaff(some);

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        final PxlDataException exception = assertThrows(PxlDataException.class, () ->
                pxl.exportExcel()
                        .workbook(workbook)
                        .override(noValidationOption())
                        .toStream(outputStream));

        assertThat(exception.getMessage().toUpperCase(Locale.ROOT)).contains("EMPLOYEES");
    }

    // ------------------------------------------------------------------
    // export: mixing workbook() and sheet() (mutually exclusive) -> exception
    // ------------------------------------------------------------------

    @Test
    public void exportBuilder_workbookAndSheetMixed_throws() {
        final GroupedWorkbook workbook = new GroupedWorkbook();
        final List<Employee> some = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, null, "Engineering"));

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        // Specifying both workbook(Object) and sheet(...) -> exception
        assertThrows(PxlArgumentException.class, () ->
                pxl.exportExcel()
                        .workbook(workbook)
                        .sheet(Employee.class, some, "People")
                        .toStream(outputStream));
    }

    @Test
    public void exportSampleBuilder_workbookAndSheetMixed_throws() {
        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        // Specifying both workbook(Class) and sheet(...) -> exception
        assertThrows(PxlArgumentException.class, () ->
                pxl.exportSampleExcel()
                        .workbook(AllTypesWorkbook.class)
                        .sheet(Employee.class, "People")
                        .toStream(outputStream));
    }

    // ------------------------------------------------------------------
    // export: neither workbook() nor sheet() specified -> exception
    // ------------------------------------------------------------------

    @Test
    public void exportBuilder_nothingSpecified_throws() {
        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        // Neither workbook(...) nor sheet(...) specified -> exception
        assertThrows(PxlArgumentException.class, () ->
                pxl.exportExcel()
                        .override(noValidationOption())
                        .toStream(outputStream));
    }

    @Test
    public void exportSampleBuilder_nothingSpecified_throws() {
        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        // Neither workbook(...) nor sheet(...) specified -> exception
        assertThrows(PxlArgumentException.class, () ->
                pxl.exportSampleExcel()
                        .toStream(outputStream));
    }

    // ------------------------------------------------------------------
    // CSV import: non-numeric value in a numeric column
    // ------------------------------------------------------------------

    @Test
    public void importCsv_invalidNumber_throws() {
        final String csv =
                "Name,Age,Salary,Active,HireDate,Grade,Department\n" +
                        "Alice,notanumber,50000,yes,2020-01-15,A,Engineering\n";

        assertThrows(PxlCellCodecException.class, () ->
                pxl.importCsv()
                        .sheet(Employee.class)
                        .fromStream("Employees", stream(csv)));
    }

    // ------------------------------------------------------------------
    // CSV import: invalid enum value
    // ------------------------------------------------------------------

    @Test
    public void importCsv_invalidEnum_throws() {
        final String csv =
                "Name,Age,Salary,Active,HireDate,Grade,Department\n" +
                        "Alice,30,50000,yes,2020-01-15,Z,Engineering\n";     // Z is not in Grade

        assertThrows(PxlCellCodecException.class, () ->
                pxl.importCsv()
                        .sheet(Employee.class)
                        .fromStream("Employees", stream(csv)));
    }

    // ------------------------------------------------------------------
    // import: wrong password
    // ------------------------------------------------------------------

    @Test
    public void importExcel_wrongPassword_throws() throws Exception {
        final Employee alice = Fixtures.employee("Alice", 30, "50000", true, null, null, "Engineering");

        final PxlExportWorkbookOption exportOption = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportPassword("secret")
                .build();
        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet(Employee.class, Arrays.asList(alice), "People")
                .override(exportOption)
                .toFile(excelFile);

        final PxlImportWorkbookOption importOption = PxlImportWorkbookOption.builder()
                .importPassword("wrong")
                .build();

        assertThrows(PxlIOException.class, () ->
                pxl.importExcel()
                        .override(importOption)
                        .sheet(Employee.class, Arrays.asList("People"))
                        .fromFile(excelFile));
    }

    // ------------------------------------------------------------------
    // Stream reader: formula cells cannot be evaluated
    // ------------------------------------------------------------------

    @Test
    public void importStreamReader_formulaCell_throws() throws Exception {
        final FormulaRow row = new FormulaRow();
        row.setLabel("calc");
        row.setFormula("=2+3");

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet(FormulaRow.class, Arrays.asList(row), "Formula")
                .override(noValidationOption())
                .toFile(excelFile);

        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .importHeaderRowIndex(1)
                .importFirstDataRowIndex(2)
                .build();
        final PxlImportWorkbookOption workbookOption = PxlImportWorkbookOption.builder()
                .importUsingStreamReader(true)
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();

        assertThrows(PxlCellCodecException.class, () ->
                pxl.importExcel()
                        .override(workbookOption)
                        .sheet(FormulaRow.class, Arrays.asList("Formula"))
                        .fromFile(excelFile));
    }

    // ------------------------------------------------------------------
    // import: Bean Validation violation (required value missing) - also validated on import
    // ------------------------------------------------------------------

    @Test
    public void importValidation_blankRequiredField_throws() {
        // CSV with an empty Name -> @NotBlank violation (importDataValidation default true)
        final String csv = "Name,Age\n,30\n";

        assertThrows(PxlValidationException.class, () ->
                pxl.importCsv()
                        .sheet(ValidatedRow.class)
                        .fromStream("Rows", stream(csv)));
    }

    @Test
    public void importValidation_nullRequiredField_throwsPxlValidation() {
        // The Age cell is empty -> Integer age = null -> @NotNull violation (importDataValidation default true).
        // Name is filled to avoid empty-row skipping (isIgnorableRow).
        final String csv = "Name,Age\nAlice,\n";

        final PxlValidationException ex = assertThrows(PxlValidationException.class, () ->
                pxl.importCsv()
                        .sheet(ValidatedRow.class)
                        .fromStream("Rows", stream(csv)));
        // Verify the constraint message for the violated field (Age) is included.
        assertTrue(ex.getMessage().contains("Age"));
    }

    @Test
    public void importValidation_emptyCollectionField_throwsPxlValidation() {
        // The Tags cell is empty -> List<String> tags = null -> @NotEmpty violation (importDataValidation default true).
        final String csv = "Name,Tags\nAlice,\n";

        final PxlValidationException ex = assertThrows(PxlValidationException.class, () ->
                pxl.importCsv()
                        .sheet(RequiredTagsRow.class)
                        .fromStream("Rows", stream(csv)));
        assertTrue(ex.getMessage().contains("Tags"));
    }

    // ------------------------------------------------------------------
    // codec value conversion errors
    // ------------------------------------------------------------------

    @Test
    public void importBoolean_invalid_throws() {
        assertImportThrows("Bool", "maybe");
    }

    @Test
    public void importBigDecimal_invalid_throws() {
        assertImportThrows("Dec", "abc");
    }

    @Test
    public void importBigInteger_invalid_throws() {
        assertImportThrows("Int", "12.5");   // decimal point for BigInteger -> format error
    }

    @Test
    public void importLocalDate_invalid_throws() {
        assertImportThrows("Date", "not-a-date");
    }

    @Test
    public void importDuration_invalid_throws() {
        assertImportThrows("Dur", "xyz");
    }

    @Test
    public void importPeriod_invalid_throws() {
        assertImportThrows("Per", "xyz");
    }

    @Test
    public void importInteger_outOfRange_throws() {
        assertImportThrows("Num", "9999999999");   // out of int range
    }

    @Test
    public void importByte_outOfRange_throws() {
        assertImportThrows("Small", "300");        // out of byte range
    }

    // ------------------------------------------------------------------
    // Per-codec defensive branches: every scalar/temporal codec must reject an unsupported (ERROR) cell
    // and an unparseable string. These exercise the switch default and the string-parse catch of each codec.
    // (char is omitted: it takes the first character of any non-empty string and never fails; boolean's two
    //  primitive/wrapper forms share one codec, so only the wrapper is listed.)
    // ------------------------------------------------------------------

    private static final String[] SCALAR_TEMPORAL_HEADERS = {
            "PrimByte", "WrapByte", "PrimShort", "WrapShort", "PrimInt", "WrapInt", "PrimLong", "WrapLong",
            "PrimDouble", "WrapDouble", "PrimFloat", "WrapFloat", "WrapBool",
            "BigInt", "BigDec", "JavaDate", "LocalDate", "LocalTime", "LocalDateTime",
            "ZonedDateTime", "OffsetTime", "OffsetDateTime", "Duration", "Period"};

    // Sheet bytes with 1 header row + 1 ERROR data cell (#DIV/0!)
    private static byte[] sheetWithErrorCell(final String header) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final Sheet sheet = workbook.createSheet("Data");
            sheet.createRow(0).createCell(0).setCellValue(header);
            sheet.createRow(1).createCell(0).setCellErrorValue(FormulaError.DIV0.getCode());
            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    @Test
    public void scalarTemporalTypes_errorCell_throwCellCodec() {
        // An ERROR cell hits each codec's switch default (unsupported cell type).
        for (final String header : SCALAR_TEMPORAL_HEADERS) {
            assertThrows(PxlCellCodecException.class,
                    () -> importTyped(sheetWithErrorCell(header), AllTypesRow.class),
                    header + " column: an unsupported (ERROR) cell must be rejected with PxlCellCodecException");
        }
    }

    @Test
    public void scalarTemporalEnum_invalidString_throwCellCodec() {
        // An unparseable string hits each codec's string-parse catch (NumberFormatException / DateTimeParseException /
        // unknown enum value). "x" is invalid for every numeric, boolean, big-number, temporal, and enum type.
        final List<String> headers = new ArrayList<>(Arrays.asList(SCALAR_TEMPORAL_HEADERS));
        headers.add("Grade");
        headers.add("Category");
        for (final String header : headers) {
            assertThrows(PxlCellCodecException.class,
                    () -> importTyped(sheetWithValue(header, "x"), AllTypesRow.class),
                    header + " column: an invalid string must be rejected with PxlCellCodecException");
        }
    }

    @Test
    public void customObjectAndCollection_invalidString_throw() {
        // Point ("x".split(",") -> parseInt("x") NFE), Money ("x".split(ws)[1] out of bounds), and a
        // collection element ("x") route their invalid input through the object/collection codec and are rejected.
        assertThrows(PxlCellCodecException.class, () -> importTyped(sheetWithValue("Point", "x"), AllTypesRow.class));
        assertThrows(PxlCellCodecException.class, () -> importTyped(sheetWithValue("Money", "x"), AllTypesRow.class));
        assertThrows(PxlCellCodecException.class, () -> importTyped(sheetWithValue("IntList", "10;x;30"), AllTypesRow.class));
    }

    // ------------------------------------------------------------------
    // Per-codec defensive branch (export side): an invalid exportSample makes each scalar/temporal codec's
    // buildXCell reject the string ("x" is unparseable as any numeric/temporal type). The sample value is
    // overridden per field via PxlExportColumnOption, so a single fixture (AllTypesRow) covers every codec.
    // (The null/blank and unsupported-object branches of buildXCell are intercepted by the resolver before
    //  dispatch, so they are unreachable defensive code and are not asserted here.)
    // ------------------------------------------------------------------

    // Java field names (not column names) of AllTypesRow whose buildXCell parses a String sample.
    private static final String[] NUMERIC_TEMPORAL_FIELD_NAMES = {
            "primByte", "wrapByte", "primShort", "wrapShort", "primInt", "wrapInt", "primLong", "wrapLong",
            "primDouble", "wrapDouble", "primFloat", "wrapFloat", "bigInt", "bigDec",
            "javaDate", "localDate", "localTime", "localDateTime", "zonedDateTime", "offsetTime", "offsetDateTime",
            "duration", "period"};

    @Test
    public void scalarTemporalTypes_invalidExportSample_throwCellCodec() throws PxlNullPointerException {
        for (final String fieldName : NUMERIC_TEMPORAL_FIELD_NAMES) {
            final PxlExportColumnOption columnOption = PxlExportColumnOption.builder()
                    .fieldName(fieldName)
                    .exportSample("x")
                    .build();
            final PxlExportSheetOption sheetOption = PxlExportSheetOption.builder()
                    .exportColumnOptions(Arrays.asList(columnOption))
                    .build();
            final PxlExportWorkbookOption workbookOption = PxlExportWorkbookOption.builder()
                    .exportSheetOptions(Arrays.asList(sheetOption))
                    .build();

            assertThrows(PxlCellCodecException.class,
                    () -> pxl.exportSampleExcel()
                            .sheet(AllTypesRow.class, "Sample")
                            .override(workbookOption)
                            .toWorkbook(),
                    fieldName + " field: an invalid exportSample must be rejected with PxlCellCodecException");
        }
    }

    @Test
    public void importErrorCell_throws() throws Exception {
        // Reading an ERROR cell (#DIV/0!) into a numeric column raises an "unsupported cell type" exception
        final byte[] bytes;
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final Sheet sheet = workbook.createSheet("Data");
            sheet.createRow(0).createCell(0).setCellValue("Num");
            final Cell cell = sheet.createRow(1).createCell(0);
            cell.setCellErrorValue(FormulaError.DIV0.getCode());
            workbook.write(outputStream);
            bytes = outputStream.toByteArray();
        }

        assertThrows(PxlCellCodecException.class, () -> importTyped(bytes, TypedRow.class));
    }

    // ------------------------------------------------------------------
    // Structural errors
    // ------------------------------------------------------------------

    @Test
    public void importRowClass_noDefaultConstructor_throws() {
        assertThrows(PxlReflectionException.class, () ->
                importTyped(sheetWithValue("Name", "Alice"), NoDefaultCtorRow.class));
    }

    @Test
    public void importRowIndex_wrongType_throws() {
        assertThrows(PxlArgumentException.class, () ->
                importTyped(sheetWithValue("Name", "Alice"), BadRowIndexRow.class));
    }

    @Test
    public void importColumn_unsupportedType_throws() {
        assertThrows(PxlArgumentException.class, () ->
                importTyped(sheetWithValue("U", "x"), UnsupportedTypeRow.class));
    }

    @Test
    public void exportGrouping_fieldTypo_throws() {
        final GroupingTypoWorkbook workbook = new GroupingTypoWorkbook();
        workbook.setWorkbookName("G");
        workbook.setEmployees(Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, Grade.A, "Engineering")));

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        assertThrows(PxlArgumentException.class, () ->
                pxl.exportExcel()
                        .workbook(workbook)
                        .override(noValidationOption())
                        .toStream(outputStream));
    }

    // ------------------------------------------------------------------
    // import: @PxlByteSize constraint violation (byte length exceeded)
    // ------------------------------------------------------------------

    @Test
    public void importByteSize_overMax_throws() {
        // "abcdef" = 6 bytes > max(5) -> @PxlByteSize violation (importDataValidation default true)
        final String csv = "Code\nabcdef\n";

        assertThrows(PxlValidationException.class, () ->
                pxl.importCsv()
                        .sheet(ByteSizeRow.class)
                        .fromStream("Rows", stream(csv)));
    }

    // ------------------------------------------------------------------
    // writeToStream: encryption requested + unsupported workbook type -> fail-fast (issue N1)
    // ------------------------------------------------------------------

    @Test
    public void writeToStream_encryptUnsupportedWorkbookType_throws() throws Exception {
        // Requesting encryption on a non-XSSF/SXSSF/HSSF workbook (StreamingWorkbook) must
        // fail with PxlArgumentException rather than silently emitting empty output.
        final byte[] xlsx;
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            workbook.createSheet("Data").createRow(0).createCell(0).setCellValue("x");
            workbook.write(outputStream);
            xlsx = outputStream.toByteArray();
        }

        try (Workbook streaming = StreamingReader.builder().open(new ByteArrayInputStream(xlsx))) {
            final ByteArrayOutputStream out = new ByteArrayOutputStream();
            assertThrows(PxlArgumentException.class, () ->
                    PxlWorkbookUtils.writeToStream(streaming, out, "secret"));
        }
    }

    // ------------------------------------------------------------------
    // PxlAssertSupport: Validate-like API -> null is PxlNullPointerException, other invalid (empty/blank) is PxlArgumentException.
    // The second String argument is inserted into the standard message as the 'parameter name'.
    // ------------------------------------------------------------------

    @Test
    public void validationSupport_notNull_nullThrowsPxlNullPointer_elseReturns() throws Exception {
        assertThrows(PxlNullPointerException.class, () -> PxlAssertSupport.notNull(null));

        // The parameter name is included in the message.
        final PxlNullPointerException npe = assertThrows(PxlNullPointerException.class,
                () -> PxlAssertSupport.notNull(null, "rowClass"));
        assertTrue(npe.getMessage().contains("rowClass"));

        // On success, returns the validated argument as-is.
        final String value = "ok";
        assertSame(value, PxlAssertSupport.notNull(value, "value"));
    }

    @Test
    public void validationSupport_notEmpty_nullIsNullPointer_emptyIsArgument() throws PxlNullPointerException, PxlArgumentException {
        // null -> PxlNullPointerException (array/collection/map/string overloads)
        assertThrows(PxlNullPointerException.class, () -> PxlAssertSupport.notEmpty((String[]) null, "arr"));
        assertThrows(PxlNullPointerException.class, () -> PxlAssertSupport.notEmpty((List<String>) null, "list"));
        assertThrows(PxlNullPointerException.class, () -> PxlAssertSupport.notEmpty((Map<String, String>) null, "map"));
        assertThrows(PxlNullPointerException.class, () -> PxlAssertSupport.notEmpty((CharSequence) null, "chars"));

        // empty -> PxlArgumentException, parameter name included in the message
        assertThrows(PxlArgumentException.class, () -> PxlAssertSupport.notEmpty(new String[0], "arr"));
        final PxlArgumentException emptyList = assertThrows(PxlArgumentException.class,
                () -> PxlAssertSupport.notEmpty(new ArrayList<String>(), "candidateSheetNames"));
        assertTrue(emptyList.getMessage().contains("candidateSheetNames"));
        assertThrows(PxlArgumentException.class, () -> PxlAssertSupport.notEmpty(new HashMap<String, String>(), "map"));
        assertThrows(PxlArgumentException.class, () -> PxlAssertSupport.notEmpty("", "chars"));

        // The no-arg overload without a parameter name behaves the same (generic default message).
        assertThrows(PxlNullPointerException.class, () -> PxlAssertSupport.notEmpty((String[]) null));
        assertThrows(PxlNullPointerException.class, () -> PxlAssertSupport.notEmpty((List<String>) null));
        assertThrows(PxlNullPointerException.class, () -> PxlAssertSupport.notEmpty((Map<String, String>) null));
        assertThrows(PxlNullPointerException.class, () -> PxlAssertSupport.notEmpty((CharSequence) null));
        assertThrows(PxlArgumentException.class, () -> PxlAssertSupport.notEmpty(new String[0]));
        assertThrows(PxlArgumentException.class, () -> PxlAssertSupport.notEmpty(new ArrayList<String>()));
        assertThrows(PxlArgumentException.class, () -> PxlAssertSupport.notEmpty(new HashMap<String, String>()));
        assertThrows(PxlArgumentException.class, () -> PxlAssertSupport.notEmpty(""));

        // On success, returns the validated argument as-is.
        final List<String> nonEmpty = Arrays.asList("a");
        assertSame(nonEmpty, PxlAssertSupport.notEmpty(nonEmpty));
    }

    @Test
    public void validationSupport_notBlank_nullIsNullPointer_blankIsArgument() throws Exception {
        assertThrows(PxlNullPointerException.class, () -> PxlAssertSupport.notBlank(null, "sheetName"));

        final PxlArgumentException blankEx = assertThrows(PxlArgumentException.class,
                () -> PxlAssertSupport.notBlank("   ", "sheetName"));
        assertTrue(blankEx.getMessage().contains("sheetName"));
        assertThrows(PxlArgumentException.class, () -> PxlAssertSupport.notBlank("", "sheetName"));

        // On success, returns the validated argument as-is.
        assertSame("v", PxlAssertSupport.notBlank("v", "sheetName"));

        // The no-arg overload without a parameter name behaves the same (generic default message).
        assertThrows(PxlNullPointerException.class, () -> PxlAssertSupport.notBlank(null));
        assertThrows(PxlArgumentException.class, () -> PxlAssertSupport.notBlank("   "));
        assertThrows(PxlArgumentException.class, () -> PxlAssertSupport.notBlank(""));
        assertSame("v", PxlAssertSupport.notBlank("v"));
    }

    @Test
    public void validationSupport_isTrue_falseThrowsArgument() throws Exception {
        assertThrows(PxlArgumentException.class, () -> PxlAssertSupport.isTrue(false, "The condition was not met."));

        // true passes (no exception)
        PxlAssertSupport.isTrue(true, "ok");

        // The no-arg overload without a message behaves the same (generic default message).
        assertThrows(PxlArgumentException.class, () -> PxlAssertSupport.isTrue(false));
        PxlAssertSupport.isTrue(true);
    }

    @Test
    public void validationSupport_exceptionSupplier_throwsSuppliedException() throws Exception {
        // On failure, throws the supplied exception as-is (need not be a Pxl exception).
        final IllegalArgumentException iae = assertThrows(IllegalArgumentException.class, () ->
                PxlAssertSupport.notEmpty(new ArrayList<String>(),
                        () -> new IllegalArgumentException("At least one tag is required.")));
        assertTrue(iae.getMessage().contains("At least one tag is required."));

        // null / blank / false all throw the supplied exception the same way.
        assertThrows(IllegalArgumentException.class, () ->
                PxlAssertSupport.notNull(null, () -> new IllegalArgumentException("required argument")));
        assertThrows(IllegalArgumentException.class, () ->
                PxlAssertSupport.notBlank("   ", () -> new IllegalArgumentException("blank not allowed")));
        assertThrows(IllegalStateException.class, () ->
                PxlAssertSupport.isTrue(false, () -> new IllegalStateException("condition violated")));

        // On success, returns the validated argument as-is (no exception).
        final List<String> tags = Arrays.asList("a", "b");
        assertSame(tags, PxlAssertSupport.notEmpty(tags, () -> new IllegalArgumentException("x")));
    }

    // ------------------------------------------------------------------
    // Builder config validation (PxlAssertSupport applied): null -> PxlNullPointerException, empty/blank -> PxlArgumentException
    // ------------------------------------------------------------------

    @Test
    public void importExcelBuilder_invalidArgs_throwPxlExceptions() {
        // workbook / rowClass / collectionClass null -> PxlNullPointerException
        assertThrows(PxlNullPointerException.class, () -> pxl.importExcel()
                .workbook(null));
        assertThrows(PxlNullPointerException.class, () -> pxl.importExcel()
                .sheet(null, Arrays.asList("Any")));
        assertThrows(PxlNullPointerException.class, () -> pxl.importExcel()
                .sheet(null, List.class, Arrays.asList("Any")));
        assertThrows(PxlNullPointerException.class, () -> pxl.importExcel()
                .sheet(Employee.class, null, Arrays.asList("Any")));

        // candidateSheetNames array null -> PxlNullPointerException
        assertThrows(PxlNullPointerException.class, () -> pxl.importExcel()
                .sheet(Employee.class, (String[]) null));

        // candidateSheetNames empty (@NotEmpty) -> PxlArgumentException
        assertThrows(PxlArgumentException.class, () -> pxl.importExcel()
                .sheet(Employee.class));
        assertThrows(PxlArgumentException.class, () -> pxl.importExcel()
                .sheet(Employee.class, new ArrayList<String>()));
    }

    @Test
    public void importCsvBuilder_nullArgs_throwPxlNullPointer() {
        assertThrows(PxlNullPointerException.class, () -> pxl.importCsv()
                .workbook(null));
        assertThrows(PxlNullPointerException.class, () -> pxl.importCsv()
                .sheet(null));
        assertThrows(PxlNullPointerException.class, () -> pxl.importCsv()
                .sheet(null, List.class));
        assertThrows(PxlNullPointerException.class, () -> pxl.importCsv()
                .sheet(Employee.class, null));
    }

    @Test
    public void exportExcelBuilder_invalidArgs_throwPxlExceptions() {
        assertThrows(PxlNullPointerException.class, () -> pxl.exportExcel()
                .workbook(null));
        assertThrows(PxlArgumentException.class, () -> pxl.exportExcel()
                .sheet(Employee.class, new ArrayList<Employee>(), "  "));
        assertThrows(PxlNullPointerException.class, () -> pxl.exportExcel()
                .sheet(Employee.class, null, "V"));
        assertThrows(PxlNullPointerException.class, () -> pxl.exportExcel()
                .sheet(null, new ArrayList<Employee>(), "V"));

        assertThrows(PxlNullPointerException.class, () -> pxl.exportSampleExcel()
                .workbook(null));
        assertThrows(PxlArgumentException.class, () -> pxl.exportSampleExcel()
                .sheet(Employee.class, "  "));
        assertThrows(PxlNullPointerException.class, () -> pxl.exportSampleExcel()
                .sheet(null, "Sample"));
    }

    @Test
    public void importBuilderTerminals_nullSource_throwPxlExceptions() {
        // Excel Source terminals: fromFile/fromStream null -> PxlNullPointerException
        assertThrows(PxlNullPointerException.class, () -> pxl.importExcel()
                .sheet(Employee.class, "Sheet1")
                .fromFile(null));
        assertThrows(PxlNullPointerException.class, () -> pxl.importExcel()
                .sheet(Employee.class, "Sheet1")
                .fromStream(null));

        // CSV Source terminals: null -> PxlNullPointerException, empty -> PxlArgumentException
        assertThrows(PxlNullPointerException.class, () -> pxl.importCsv()
                .sheet(Employee.class)
                .fromFile(null));
        assertThrows(PxlNullPointerException.class, () -> pxl.importCsv()
                .sheet(Employee.class)
                .fromFiles(null));
        assertThrows(PxlArgumentException.class, () -> pxl.importCsv()
                .sheet(Employee.class)
                .fromFiles(new ArrayList<File>()));
        assertThrows(PxlNullPointerException.class, () -> pxl.importCsv()
                .sheet(Employee.class)
                .fromStream(null, new ByteArrayInputStream(new byte[0])));
        assertThrows(PxlNullPointerException.class, () -> pxl.importCsv()
                .sheet(Employee.class)
                .fromStream("name", null));
        assertThrows(PxlNullPointerException.class, () -> pxl.importCsv()
                .sheet(Employee.class)
                .fromStreams(null, null));
    }

    @Test
    public void exportBuilderTerminals_nullDestination_throwPxlNullPointer() {
        final List<Employee> rows = new ArrayList<>();
        assertThrows(PxlNullPointerException.class, () -> pxl.exportExcel()
                .sheet(Employee.class, rows, "Sheet1")
                .toFile(null));
        assertThrows(PxlNullPointerException.class, () -> pxl.exportExcel()
                .sheet(Employee.class, rows, "Sheet1")
                .toStream(null));
    }

    @Test
    public void optionAddMethods_nullArg_throwPxlNullPointer() {
        assertThrows(PxlNullPointerException.class, () -> PxlExportWorkbookOption.builder().build().addExportSheetOption(null));
        assertThrows(PxlNullPointerException.class, () -> PxlImportWorkbookOption.builder().build().addImportSheetOption(null));
        assertThrows(PxlNullPointerException.class, () -> PxlExportSheetOption.builder().build().addExportColumnOption(null));
        assertThrows(PxlNullPointerException.class, () -> PxlImportSheetOption.builder().build().addImportColumnOption(null));
    }

    @Test
    public void workbookUtils_nullArgs_throwPxlNullPointer() {
        assertThrows(PxlNullPointerException.class, () -> PxlWorkbookUtils.openWorkbook((File) null, null, false));
        assertThrows(PxlNullPointerException.class, () -> PxlWorkbookUtils.openWorkbook((InputStream) null, null));
        assertThrows(PxlNullPointerException.class, () -> PxlWorkbookUtils.writeToStream(null, new ByteArrayOutputStream(), null));
        assertThrows(PxlNullPointerException.class, () -> PxlWorkbookUtils.writeToStream(new XSSFWorkbook(), null, null));
        assertThrows(PxlNullPointerException.class, () -> PxlWorkbookUtils.createFormulaEvaluator(null));
    }

    @Test
    public void publicUtils_missingOrUnusableArgs_throwPxlException() {
        // Each of these used to let a raw POI/JDK exception out of a public util entry point.
        assertThrows(PxlNullPointerException.class, () -> PxlSheetUtils.cloneSheet(null, 0, "Cloned"));
        assertThrows(PxlNullPointerException.class, () -> PxlMiscUtils.convertColumnStringToColumnIndex(null));
        assertThrows(PxlNullPointerException.class, () -> PxlMiscUtils.convertCellReferenceStringToIndexes(null));
        // A pictures-per-row count of 0 divided by zero while sizing the grid, even for an empty picture list.
        assertThrows(PxlArgumentException.class,
                () -> PxlCellUtils.addPicturesToCell((Sheet) null, new ArrayList<>(), 100, 100, 5, 0, 0, 0));
        assertThrows(PxlArgumentException.class,
                () -> PxlCellUtils.addPicturesToCell((Cell) null, new ArrayList<>(), 100, 100, 5, 0));
    }

    // ------------------------------------------------------------------
    // Exception type constructors: the four common forms (no-arg / message / message+cause / cause)
    // carry the detail message and cause through to Throwable, and the location-tagged forms embed the
    // sheet/row/column values into the message. These are the direct-construction paths not otherwise
    // exercised by the wrapping happy/error flows above.
    // ------------------------------------------------------------------

    @FunctionalInterface
    private interface NoArgCtor {
        Throwable make();
    }

    @FunctionalInterface
    private interface MessageCtor {
        Throwable make(String message);
    }

    @FunctionalInterface
    private interface MessageCauseCtor {
        Throwable make(String message, Throwable cause);
    }

    @FunctionalInterface
    private interface CauseCtor {
        Throwable make(Throwable cause);
    }

    // Exercises the four common constructor forms shared by every Pxl exception type.
    private static void assertCommonConstructors(final NoArgCtor noArg,
                                                 final MessageCtor message,
                                                 final MessageCauseCtor messageCause,
                                                 final CauseCtor cause) {
        assertNull(noArg.make().getMessage());

        assertEquals("boom", message.make("boom").getMessage());

        final Throwable root = new IllegalStateException("root");
        final Throwable wrapped = messageCause.make("wrap", root);
        assertEquals("wrap", wrapped.getMessage());
        assertSame(root, wrapped.getCause());

        assertSame(root, cause.make(root).getCause());
    }

    @Test
    public void exceptionTypes_commonConstructors_carryMessageAndCause() {
        assertCommonConstructors(PxlArgumentException::new, PxlArgumentException::new, PxlArgumentException::new, PxlArgumentException::new);
        assertCommonConstructors(PxlNullPointerException::new, PxlNullPointerException::new, PxlNullPointerException::new, PxlNullPointerException::new);
        assertCommonConstructors(PxlCellCodecException::new, PxlCellCodecException::new, PxlCellCodecException::new, PxlCellCodecException::new);
        assertCommonConstructors(PxlValidationException::new, PxlValidationException::new, PxlValidationException::new, PxlValidationException::new);
        assertCommonConstructors(PxlDataException::new, PxlDataException::new, PxlDataException::new, PxlDataException::new);
        assertCommonConstructors(PxlReflectionException::new, PxlReflectionException::new, PxlReflectionException::new, PxlReflectionException::new);
        assertCommonConstructors(PxlI18nException::new, PxlI18nException::new, PxlI18nException::new, PxlI18nException::new);
        assertCommonConstructors(PxlIOException::new, PxlIOException::new, PxlIOException::new, PxlIOException::new);
        assertCommonConstructors(PxlRuntimeException::new, PxlRuntimeException::new, PxlRuntimeException::new, PxlRuntimeException::new);
        assertCommonConstructors(PxlSystemException::new, PxlSystemException::new, PxlSystemException::new, PxlSystemException::new);
    }

    @Test
    public void exceptionTypes_taggedConstructors_embedLocationValues() {
        // message form: the sheet name, one-based row (index 4 -> "5"), column name, and message all appear.
        final PxlValidationException byName = new PxlValidationException("Sheet1", 4, "Age", null, "boom");
        assertTrue(byName.getMessage().contains("Sheet1"));
        assertTrue(byName.getMessage().contains("Age"));
        assertTrue(byName.getMessage().contains("boom"));
        assertTrue(byName.getMessage().contains("5"));

        // column-index form: index 2 renders as the spreadsheet column letter "C".
        final PxlCellCodecException byIndex = new PxlCellCodecException("Data", 0, null, 2, "bad cell");
        assertTrue(byIndex.getMessage().contains("Data"));
        assertTrue(byIndex.getMessage().contains("C"));
        assertTrue(byIndex.getMessage().contains("bad cell"));

        // message+cause form keeps the cause.
        final Throwable root = new IllegalStateException("root");
        final PxlCellCodecException withCause = new PxlCellCodecException("Data", 1, "Col", null, "oops", root);
        assertTrue(withCause.getMessage().contains("Col"));
        assertTrue(withCause.getMessage().contains("oops"));
        assertSame(root, withCause.getCause());

        // cause-derived form derives the message from the cause and keeps the cause.
        final PxlValidationException derived = new PxlValidationException("Data", 1, "Col", null, root);
        assertTrue(derived.getMessage().contains("Col"));
        assertSame(root, derived.getCause());
    }

}
