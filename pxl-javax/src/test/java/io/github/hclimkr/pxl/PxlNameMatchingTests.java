package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.exception.PxlDataException;
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;
import io.github.hclimkr.pxl.tcdata.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInfo;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Column/sheet name matching rule tests.
 * <p>
 * Names are matched whitespace-insensitively; a column header is matched case-sensitively while a sheet name is not.
 * Verifies array (alias), numeric-header, and enum-value matching rules, the name a column is written under (the
 * first alias, or the field name when none is declared), plus behavior when a required/optional column header is
 * missing.
 */
public class PxlNameMatchingTests {

    private static Pxl pxl;

    @BeforeAll
    public static void setUpBeforeClass() {
        pxl = new Pxl();
    }

    private TestInfo testInfo;

    @BeforeEach
    public void bindTestInfo(final TestInfo testInfo) {
        this.testInfo = testInfo;
    }

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

    // ------------------------------------------------------------------
    // Column name matching: whitespace-insensitive / case-sensitive / alias / numeric header
    // ------------------------------------------------------------------

    @Test
    public void columnName_headerWithWhitespace_matches() throws Exception {
        // Header "Full Name" (with space) == column "FullName"
        final byte[] bytes = stringSheet("M", new String[]{"Full Name", "Id"}, new String[][]{{"Alice", "1"}});
        assertThat(importList(bytes, "M", NameMatchRow.class).get(0).getName()).isEqualTo("Alice");
    }

    @Test
    public void columnName_headerDifferentCase_doesNotMatch() throws Exception {
        // Header "fullname" (lowercase) != column "FullName" -> no match (optional) -> null
        final byte[] bytes = stringSheet("M", new String[]{"fullname", "Id"}, new String[][]{{"Alice", "1"}});
        final NameMatchRow row = importList(bytes, "M", NameMatchRow.class).get(0);
        assertThat(row.getName()).isNull();
        assertThat(row.getId()).isEqualTo("1");
    }

    @Test
    public void columnName_alias_anyOfArrayMatches() throws Exception {
        // @PxlColumn(name={"FullName","Name","성명"}) - matches even when the header is the alias "Name"
        final byte[] bytes = stringSheet("Alias", new String[]{"Name", "Age"}, new String[][]{{"Alice", "30"}});
        final AliasRow row = importList(bytes, "Alias", AliasRow.class).get(0);
        assertThat(row.getName()).isEqualTo("Alice");
        assertThat(row.getAge()).isEqualTo(30);
    }

    @Test
    public void columnName_alias_exportsUnderTheFirstName() throws Exception {
        // Import accepts any of the aliases, but a header can only be written as one of them: the first.
        final AliasRow row = new AliasRow();
        row.setName("Alice");
        row.setAge(30);

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet(AliasRow.class, Arrays.asList(row), "Alias")
                .override(PxlExportWorkbookOption.builder().exportDataValidation(false).build())
                .toFile(excelFile);

        try (Workbook poi = WorkbookFactory.create(excelFile)) {
            final Row header = poi.getSheet("Alias").getRow(0);
            assertThat(Arrays.asList(header.getCell(0).getStringCellValue(), header.getCell(1).getStringCellValue()))
                    .as("the remaining aliases are not written as headers")
                    .containsExactlyInAnyOrder("FullName", "Age");
        }
    }

    @Test
    public void columnName_notDeclared_usesFieldName() throws Exception {
        // With name left at {}, the field name is the header written and the header matched.
        final FieldNameColumnRow row = new FieldNameColumnRow();
        row.setCode("A-1");
        row.setAmount(7);

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet(FieldNameColumnRow.class, Arrays.asList(row), "Fields")
                .override(PxlExportWorkbookOption.builder().exportDataValidation(false).build())
                .toFile(excelFile);

        try (Workbook poi = WorkbookFactory.create(excelFile)) {
            final Row header = poi.getSheet("Fields").getRow(0);
            assertThat(Arrays.asList(header.getCell(0).getStringCellValue(), header.getCell(1).getStringCellValue()))
                    .containsExactlyInAnyOrder("code", "amount");
        }

        final FieldNameColumnRow imported = pxl.importExcel()
                .sheet(FieldNameColumnRow.class, Arrays.asList("Fields"))
                .fromFile(excelFile)
                .get(0);

        assertThat(imported.getCode()).isEqualTo("A-1");
        assertThat(imported.getAmount()).isEqualTo(7);
    }

    @Test
    public void columnName_numericHeaderCell_matches() throws Exception {
        // Numeric header cell 2024 == column name "2024"
        final byte[] bytes = sheet("H", s -> {
            s.createRow(0).createCell(0).setCellValue(2024);
            s.createRow(1).createCell(0).setCellValue("value");
        });
        assertThat(importList(bytes, "H", NumericHeaderRow.class).get(0).getYear()).isEqualTo("value");
    }

    // ------------------------------------------------------------------
    // Sheet name matching: alias / case-insensitive
    // ------------------------------------------------------------------

    @Test
    public void sheetName_differentCase_matches() throws Exception {
        // Actual sheet "Data" == candidate name "DATA" -> matched ignoring case
        // (a sheet name is not always typed by hand: a CSV sheet is named after its file)
        final byte[] bytes = stringSheet("Data", new String[]{"Full Name", "Id"}, new String[][]{{"Alice", "1"}});
        assertThat(importList(bytes, "DATA", NameMatchRow.class).get(0).getName()).isEqualTo("Alice");
    }

    @Test
    public void sheetName_workbookForm_differentCase_matches() throws Exception {
        // The workbook form matches the same way: the actual sheet "EMPLOYEE" binds the @PxlSheet alias "Employee"
        final File excelFile = TestPaths.exportFile(testInfo);
        final List<Employee> employees = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, Grade.A, "Engineering"));
        pxl.exportExcel()
                .sheet(Employee.class, employees, "EMPLOYEE")
                .override(PxlExportWorkbookOption.builder().exportDataValidation(false).build())
                .toFile(excelFile);

        final AliasSheetWorkbook imported = pxl.importExcel()
                .workbook(AliasSheetWorkbook.class)
                .fromFile(excelFile);

        assertThat(imported.getData()).extracting(Employee::getName).containsExactly("Alice");
    }

    @Test
    public void sheetName_alias_anyOfArrayMatches() throws Exception {
        // @PxlSheet(name={"Crew","Employee","직원"}) - matches even when the actual sheet name is "Employee"
        final File excelFile = TestPaths.exportFile(testInfo);
        final List<Employee> employees = Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, Grade.A, "Engineering"),
                Fixtures.employee("Bob", 42, "72000", false, null, Grade.B, "Sales"));
        pxl.exportExcel()
                .sheet(Employee.class, employees, "Employee")
                .override(PxlExportWorkbookOption.builder().exportDataValidation(false).build())
                .toFile(excelFile);

        final AliasSheetWorkbook imported = pxl.importExcel()
                .workbookName("Aliased")
                .workbook(AliasSheetWorkbook.class)
                .fromFile(excelFile);

        assertThat(imported.getData()).extracting(Employee::getName).containsExactly("Alice", "Bob");
    }

    // ------------------------------------------------------------------
    // Enum value matching: case- and whitespace-insensitive
    // ------------------------------------------------------------------

    @Test
    public void enumValue_caseAndWhitespaceInsensitive_matches() throws Exception {
        // "food & beverage" (lowercase) == Category.FOOD("Food & Beverage")
        final byte[] bytes = stringSheet("E", new String[]{"Cat"}, new String[][]{{"food & beverage"}});
        assertThat(importList(bytes, "E", EnumCaseRow.class).get(0).getCat()).isEqualTo(Category.FOOD);
    }

    // ------------------------------------------------------------------
    // Missing required / optional column header
    // ------------------------------------------------------------------

    @Test
    public void requiredColumn_headerMissing_throws() throws Exception {
        // The "Req" (@NotBlank) header is missing
        final byte[] bytes = stringSheet("R", new String[]{"Opt"}, new String[][]{{"x"}});
        assertThrows(PxlDataException.class, () -> importList(bytes, "R", RequiredColRow.class));
    }

    @Test
    public void requiredColumn_headerMissing_validationDisabled_stillThrows() throws Exception {
        // The required-column existence check runs independently of importDataValidation:
        // the "Req" (@NotBlank) header is missing, so import throws even with data validation disabled.
        final byte[] bytes = stringSheet("R", new String[]{"Opt"}, new String[][]{{"x"}});
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importDataValidation(false)
                .build();
        assertThrows(PxlDataException.class, () -> pxl.importExcel()
                .override(option)
                .sheet(RequiredColRow.class, Arrays.asList("R"))
                .fromStream(new ByteArrayInputStream(bytes)));
    }

    @Test
    public void optionalColumn_headerMissing_isNull() throws Exception {
        // The "Opt" (optional) header is missing -> null, no exception
        final byte[] bytes = stringSheet("R", new String[]{"Req"}, new String[][]{{"r"}});
        final RequiredColRow row = importList(bytes, "R", RequiredColRow.class).get(0);
        assertThat(row.getReq()).isEqualTo("r");
        assertThat(row.getOpt()).isNull();
    }
}
