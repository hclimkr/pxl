package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.exception.PxlDataException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.exception.PxlValidationException;
import io.github.hclimkr.pxl.option.*;
import io.github.hclimkr.pxl.tcdata.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInfo;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.*;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Workbook/option-level behavior tests - importDataValidation, sheet/column exportOrder, superclass override,
 * stream-reader cache buffer, SXSSF window, conditional sheets, column export on/off, enum dropdown, importColumnRange.
 */
public class PxlWorkbookOptionTests {

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

    private static byte[] buildStringSheet(final String sheetName, final String[] headers, final String[][] dataRows) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final Sheet sheet = workbook.createSheet(sheetName);
            final Row header = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i).setCellValue(headers[i]);
            }
            for (int r = 0; r < dataRows.length; r++) {
                final Row row = sheet.createRow(r + 1);
                for (int c = 0; c < dataRows[r].length; c++) {
                    row.createCell(c).setCellValue(dataRows[r][c]);
                }
            }
            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    private static Set<String> headers(final Workbook workbook, final String sheetName) {
        final Row header = workbook.getSheet(sheetName).getRow(0);
        final Set<String> set = new HashSet<>();
        for (final Cell cell : header) {
            set.add(cell.getStringCellValue());
        }
        return set;
    }

    private static int colIndex(final Sheet sheet, final String header) {
        for (final Cell cell : sheet.getRow(0)) {
            if (header.equals(cell.getStringCellValue())) {
                return cell.getColumnIndex();
            }
        }
        throw new IllegalStateException("header not found: " + header);
    }

    private static List<Employee> twoEmployees() {
        return Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, Grade.A, "Engineering"),
                Fixtures.employee("Bob", 42, "72000", false, null, Grade.B, "Sales"));
    }

    // Returns the header strings of the given sheet's header row (0-based 0) in declaration order.
    private static List<String> headerSequence(final Sheet sheet) {
        final Row header = sheet.getRow(0);
        final List<String> headers = new ArrayList<>();
        for (final Cell cell : header) {
            headers.add(cell.getStringCellValue());
        }
        return headers;
    }

    // col0="Name", col1="Age", col2="Name" (duplicate). A duplicate header triggers an exception during matching.
    private static byte[] buildDuplicateHeaderFixture() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            final Sheet sheet = workbook.createSheet("Dup");

            final Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Name");
            header.createCell(1).setCellValue("Age");
            header.createCell(2).setCellValue("Name");      // duplicate header

            final Row data = sheet.createRow(1);
            data.createCell(0).setCellValue("Alice");
            data.createCell(1).setCellValue(30);
            data.createCell(2).setCellValue("Duplicate");

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    // ------------------------------------------------------------------
    // importDataValidation=false -> even constraint-violating data is imported without an exception
    // ------------------------------------------------------------------

    @Test
    public void importDataValidation_disabled_skipsValidation() throws Exception {
        // Name is empty (@NotBlank violation), Age 20
        final byte[] bytes = buildStringSheet("V", new String[]{"Name", "Age"}, new String[][]{{"", "20"}});

        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importDataValidation(false)
                .build();

        final List<ValidatedRow> rows = pxl.importExcel()
                .override(option)
                .sheet(ValidatedRow.class, Arrays.asList("V"))
                .fromStream(new ByteArrayInputStream(bytes));

        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).getName()).isNull();    // validation skipped -> null without an exception
        assertThat(rows.get(0).getAge()).isEqualTo(20);
    }

    @Test
    public void importDataValidation_default_blankRequiredValue_throws() throws Exception {
        // Counterpart to the disabled case: the "Name" column exists but its value is empty
        // (@NotBlank violation). With importDataValidation defaulting to true, value validation runs
        // and Excel import fails with PxlValidationException.
        final byte[] bytes = buildStringSheet("V", new String[]{"Name", "Age"}, new String[][]{{"", "20"}});

        assertThrows(PxlValidationException.class, () -> pxl.importExcel()
                .sheet(ValidatedRow.class, Arrays.asList("V"))
                .fromStream(new ByteArrayInputStream(bytes)));
    }

    // ------------------------------------------------------------------
    // Import .override(...) chain position - before or after workbook(...)/sheet(...)
    // ------------------------------------------------------------------

    @Test
    public void importOverride_afterSheetConfig_skipsValidation() throws Exception {
        // Same option as importDataValidation_disabled_skipsValidation, but chained after sheet(...)
        // on the source step - it must apply exactly like the builder-side call.
        final byte[] bytes = buildStringSheet("V", new String[]{"Name", "Age"}, new String[][]{{"", "20"}});

        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importDataValidation(false)
                .build();

        final List<ValidatedRow> rows = pxl.importExcel()
                .sheet(ValidatedRow.class, Arrays.asList("V"))
                .override(option)
                .fromStream(new ByteArrayInputStream(bytes));

        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).getName()).isNull();    // validation skipped -> null without an exception
        assertThat(rows.get(0).getAge()).isEqualTo(20);
    }

    @Test
    public void importOverride_beforeAndAfterSheetConfig_lastWins() throws Exception {
        // Both sides set an option; the one chained last (on the source step) wins,
        // so validation is skipped even though the builder-side option enabled it.
        final byte[] bytes = buildStringSheet("V", new String[]{"Name", "Age"}, new String[][]{{"", "20"}});

        final PxlImportWorkbookOption validating = PxlImportWorkbookOption.builder()
                .importDataValidation(true)
                .build();
        final PxlImportWorkbookOption notValidating = PxlImportWorkbookOption.builder()
                .importDataValidation(false)
                .build();

        final List<ValidatedRow> rows = pxl.importExcel()
                .override(validating)
                .sheet(ValidatedRow.class, Arrays.asList("V"))
                .override(notValidating)
                .fromStream(new ByteArrayInputStream(bytes));

        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).getName()).isNull();
        assertThat(rows.get(0).getAge()).isEqualTo(20);
    }

    // ------------------------------------------------------------------
    // Sheet exportOrder -> sheet order
    // ------------------------------------------------------------------

    @Test
    public void sheetExportOrder_ordersSheets() throws Exception {
        final SheetOrderWorkbook workbook = new SheetOrderWorkbook();
        workbook.setWorkbookName("Order");
        workbook.setZebra(twoEmployees());
        workbook.setApple(twoEmployees());

        final Workbook poi = pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption())
                .toWorkbook();
        try {
            // exportOrder order: A(Zebra), B(Apple)
            assertThat(poi.getSheetName(0)).isEqualTo("Zebra");
            assertThat(poi.getSheetName(1)).isEqualTo("Apple");
        } finally {
            poi.close();
        }
    }

    // ------------------------------------------------------------------
    // Sheet override on export (exportOverrideSuperClassSheet)
    // ------------------------------------------------------------------

    @Test
    public void exportOverrideSuperClassSheet_childSheetWins() throws Exception {
        final SubExportSheetWorkbook workbook = new SubExportSheetWorkbook();
        workbook.workbookName = "W";
        workbook.employees = twoEmployees();                                   // child data (Alice, Bob)
        ((SuperExportSheetWorkbook) workbook).employees = Arrays.asList(
                Fixtures.employee("Carol", 35, "68000", true, null, Grade.A, "Finance"));   // parent data

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption())
                .toFile(excelFile);

        // child overrides -> a single "Employees" sheet with only the child data
        final List<Employee> imported = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("Employees"))
                .fromFile(excelFile);

        assertThat(imported).extracting(Employee::getName).containsExactly("Alice", "Bob");
    }

    @Test
    public void exportOverrideSuperClassSheet_differentCase_childSheetWins() throws Exception {
        // The child names the same sheet in a different case ("EMPLOYEES" against the super's "Employees").
        // Names differing only in case denote one sheet - a workbook cannot hold both - so the override applies.
        final SubCaseExportSheetWorkbook workbook = new SubCaseExportSheetWorkbook();
        workbook.workbookName = "W";
        workbook.employees = twoEmployees();                                   // child data (Alice, Bob)
        ((SuperExportSheetWorkbook) workbook).employees = Arrays.asList(
                Fixtures.employee("Carol", 35, "68000", true, null, Grade.A, "Finance"));   // parent data

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption())
                .toFile(excelFile);

        // child overrides -> a single sheet with only the child data, read back with the differently cased name
        final List<Employee> imported = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("Employees"))
                .fromFile(excelFile);

        assertThat(imported).extracting(Employee::getName).containsExactly("Alice", "Bob");
    }

    // Returns the "Name" column values of the given sheet's data rows (header row 0).
    private static List<String> employeeNames(final Sheet sheet) {
        final int nameColumn = colIndex(sheet, "Name");
        final List<String> names = new ArrayList<>();
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            names.add(sheet.getRow(rowIndex).getCell(nameColumn).getStringCellValue());
        }
        return names;
    }

    // A workbook whose super field carries data too, so a suppressed sheet is suppressed by the override
    // rather than by having nothing to write.
    private static List<Employee> oneParentEmployee() {
        return Arrays.asList(Fixtures.employee("Carol", 35, "68000", true, null, Grade.A, "Finance"));
    }

    @Test
    public void exportOverrideSuperClassSheet_bySheetNameNotFieldName_childSheetWins() throws Exception {
        // The override is keyed on the sheet name, not the field name: a field named staff that lists "Employees"
        // among its candidate names still suppresses the super's employees field. The sheet it writes is named
        // after the FIRST candidate ("Crew") - matching decides the override, the first name decides the label.
        final SubAliasOverrideExportSheetWorkbook workbook = new SubAliasOverrideExportSheetWorkbook();
        workbook.workbookName = "W";
        workbook.staff = twoEmployees();                                       // child data (Alice, Bob)
        ((SuperExportSheetWorkbook) workbook).employees = oneParentEmployee();  // parent data

        final Workbook poi = pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption())
                .toWorkbook();
        try {
            assertThat(poi.getNumberOfSheets()).isEqualTo(1);
            assertThat(poi.getSheetName(0)).isEqualTo("Crew");
            assertThat(employeeNames(poi.getSheetAt(0))).containsExactly("Alice", "Bob");
        } finally {
            poi.close();
        }
    }

    @Test
    public void exportOverrideSuperClassSheet_withDifferentSheetName_doesNotOverride() throws Exception {
        // Shadowing the field is not by itself an override: this employees field names a different sheet
        // ("Staff"), so the super's "Employees" is written as well and the workbook carries both.
        final SubOtherNameOverrideExportSheetWorkbook workbook = new SubOtherNameOverrideExportSheetWorkbook();
        workbook.workbookName = "W";
        workbook.employees = twoEmployees();
        ((SuperExportSheetWorkbook) workbook).employees = oneParentEmployee();

        final Workbook poi = pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption())
                .toWorkbook();
        try {
            assertThat(poi.getNumberOfSheets()).isEqualTo(2);
            assertThat(employeeNames(poi.getSheet("Staff"))).containsExactly("Alice", "Bob");
            assertThat(employeeNames(poi.getSheet("Employees"))).containsExactly("Carol");
        } finally {
            poi.close();
        }
    }

    @Test
    public void exportOverrideSuperClassSheet_onDisabledSheet_doesNotOverride() throws Exception {
        // A sheet excluded from export claims no name, so its override never takes effect:
        // the super's employees field writes the "Employees" sheet.
        final SubDisabledOverrideExportSheetWorkbook workbook = new SubDisabledOverrideExportSheetWorkbook();
        workbook.workbookName = "W";
        workbook.employees = twoEmployees();
        ((SuperExportSheetWorkbook) workbook).employees = oneParentEmployee();

        final Workbook poi = pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption())
                .toWorkbook();
        try {
            assertThat(poi.getNumberOfSheets()).isEqualTo(1);
            assertThat(poi.getSheetName(0)).isEqualTo("Employees");
            assertThat(employeeNames(poi.getSheetAt(0))).containsExactly("Carol");
        } finally {
            poi.close();
        }
    }

    @Test
    public void exportOverrideSuperClassSheet_declaredOnSuperClass_throws() throws Exception {
        // The flag runs from a subclass toward its superclass only. Declared on the superclass it suppresses
        // nothing - the subclass field is resolved first - so both fields ask for an "Employees" sheet and the
        // export fails, where the same pair with the flag on the child (see above) succeeds.
        // The duplicate is caught while resolving the sheet metadata, before any sheet is created.
        final SubOverrideExportSheetWorkbook workbook = new SubOverrideExportSheetWorkbook();
        workbook.workbookName = "W";
        workbook.employees = twoEmployees();
        ((SuperOverrideExportSheetWorkbook) workbook).employees = oneParentEmployee();

        assertThrows(PxlDataException.class, () -> pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption())
                .toWorkbook());
    }

    // ------------------------------------------------------------------
    // Column override on export (exportOverrideSuperClassColumn)
    // ------------------------------------------------------------------

    @Test
    public void exportOverrideSuperClassColumn_childColumnWins() throws Exception {
        final DerivedColRow row = new DerivedColRow();
        row.val = "sub";                        // child field
        ((BaseColRow) row).val = "super";       // parent field

        final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        pxl.exportExcel()
                .sheet(DerivedColRow.class, Arrays.asList(row), "C")
                .override(noValidationOption())
                .toStream(outputStream);

        try (Workbook poi = WorkbookFactory.create(new ByteArrayInputStream(outputStream.toByteArray()))) {
            final Sheet sheet = poi.getSheet("C");
            // only one "Val" column, with the child field's value
            final Set<String> headerSet = headers(poi, "C");
            assertThat(headerSet).containsExactly("Val");
            assertThat(sheet.getRow(1).getCell(colIndex(sheet, "Val")).getStringCellValue()).isEqualTo("sub");
        }
    }

    // ------------------------------------------------------------------
    // Works correctly even when stream-reader cache/buffer sizes are specified
    // ------------------------------------------------------------------

    @Test
    public void streamReaderCacheBuffer_variants_work() throws Exception {
        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet(Employee.class, twoEmployees(), "People")
                .override(noValidationOption())
                .toFile(excelFile);

        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .importHeaderRowIndex(1)
                .importFirstDataRowIndex(2)
                .build();
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importUsingStreamReader(true)
                .importStreamReaderRowCacheSize(50)
                .importStreamReaderBufferSize(8192)
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();

        final List<Employee> rows = pxl.importExcel()
                .override(option)
                .sheet(Employee.class, Arrays.asList("People"))
                .fromFile(excelFile);

        assertThat(rows).extracting(Employee::getName).containsExactly("Alice", "Bob");
    }

    // ------------------------------------------------------------------
    // Works correctly even when the SXSSF row-access window size is specified
    // ------------------------------------------------------------------

    @Test
    public void sxssfRowAccessWindow_variants_work() throws Exception {
        final PxlExportWorkbookOption exportOption = PxlExportWorkbookOption.builder()
                .exportExcelEngine(PxlExcelEngine.SXSSF)
                .exportSXSSFRowAccessWindowSize(100)
                .exportDataValidation(false)
                .build();

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet(Employee.class, twoEmployees(), "People")
                .override(exportOption)
                .toFile(excelFile);

        final List<Employee> rows = pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("People"))
                .fromFile(excelFile);

        assertThat(rows).extracting(Employee::getName).containsExactly("Alice", "Bob");
    }

    // ------------------------------------------------------------------
    // exportOrder: column order follows the order strings (alphabetical)
    // ------------------------------------------------------------------

    @Test
    public void exportOrder_controlsColumnOrder() throws Exception {
        final OrderedRow row = new OrderedRow();
        row.setX("x");
        row.setY("y");
        row.setZ("z");

        final Workbook workbook = pxl.exportExcel()
                .sheet(OrderedRow.class, Arrays.asList(row), "Order")
                .override(noValidationOption())
                .toWorkbook();
        try {
            // declaration order is X,Y,Z but exportOrder gives Y(A), Z(B), X(C)
            assertThat(headerSequence(workbook.getSheet("Order"))).containsExactly("Y", "Z", "X");
        } finally {
            workbook.close();
        }
    }

    // ------------------------------------------------------------------
    // exportIfNull / exportIfEmpty: conditional sheet creation
    // ------------------------------------------------------------------

    @Test
    public void conditionalSheet_nullOrEmpty_excluded() throws Exception {
        final ConditionalWorkbook workbook = new ConditionalWorkbook();
        workbook.setKeepWhenNull(null);                 // exportIfNull=true  -> created
        workbook.setDropWhenNull(null);                 // default (false)    -> not created
        workbook.setDropWhenEmpty(new ArrayList<>());   // exportIfEmpty=false -> not created
        workbook.setKeepWhenEmpty(new ArrayList<>());   // default (true)     -> created

        final Workbook poiWorkbook = pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption())
                .toWorkbook();
        try {
            assertThat(poiWorkbook.getSheet("KeepWhenNull")).as("created even when null").isNotNull();
            assertThat(poiWorkbook.getSheet("DropWhenNull")).as("not created when null").isNull();
            assertThat(poiWorkbook.getSheet("DropWhenEmpty")).as("not created when list is empty").isNull();
            assertThat(poiWorkbook.getSheet("KeepWhenEmpty")).as("created even when list is empty").isNotNull();
        } finally {
            poiWorkbook.close();
        }
    }

    // ------------------------------------------------------------------
    // exportEnabled=false (option): exclude a specific column
    // ------------------------------------------------------------------

    @Test
    public void columnExportDisabled_excludesColumn() throws Exception {
        final Employee employee = Fixtures.employee("Alice", 30, "50000", true, null, Grade.A, "Engineering");

        final PxlExportColumnOption columnOption = PxlExportColumnOption.builder()
                .fieldName("age")
                .exportEnabled(false)
                .build();
        final PxlExportSheetOption sheetOption = PxlExportSheetOption.builder()
                .exportColumnOptions(Arrays.asList(columnOption))
                .build();
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportSheetOptions(Arrays.asList(sheetOption))
                .build();

        final Workbook workbook = pxl.exportExcel()
                .sheet(Employee.class, Arrays.asList(employee), "People")
                .override(option)
                .toWorkbook();
        try {
            final List<String> headers = headerSequence(workbook.getSheet("People"));
            assertThat(headers).contains("Name", "Salary", "Department");
            assertThat(headers).doesNotContain("Age");      // excluded by exportEnabled=false
        } finally {
            workbook.close();
        }
    }

    // ------------------------------------------------------------------
    // Nested 3-level override (workbook -> sheet -> column) on import:
    // a column-level importEnabled=false skips binding just that column
    // ------------------------------------------------------------------

    @Test
    public void nestedColumnOverride_importDisabledColumn_skipsBinding() throws Exception {
        // Export two employees (ages 30, 42) with every column present.
        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet(Employee.class, twoEmployees(), "People")
                .override(noValidationOption())
                .toFile(excelFile);

        // Build the full workbook -> sheet -> column override tree and pass it in one override(...) call.
        final PxlImportColumnOption ageColumn = PxlImportColumnOption.builder()
                .fieldName("age")                                   // matched to the age column by field name
                .importEnabled(false)                               // disable binding for this column
                .build();
        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()   // field name omitted -> wildcard (any sheet)
                .importColumnOptions(Arrays.asList(ageColumn))
                .build();
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();

        final List<Employee> rows = pxl.importExcel()
                .override(option)
                .sheet(Employee.class, Arrays.asList("People"))
                .fromFile(excelFile);

        // Other columns are still bound, but the disabled age column stays at the primitive default (0).
        assertThat(rows).extracting(Employee::getName).containsExactly("Alice", "Bob");
        assertThat(rows).extracting(Employee::getAge).containsOnly(0);
    }

    // ------------------------------------------------------------------
    // enum dropdown: data validation is created when exportDataValidation=true
    // ------------------------------------------------------------------

    @Test
    public void enumDropdown_dataValidationPresent() throws Exception {
        final Employee employee = Fixtures.employee("Alice", 30, "50000", true, null, Grade.A, "Engineering");

        // option null -> exportDataValidation defaults to true -> dropdown created on the Grade enum column
        final Workbook workbook = pxl.exportExcel()
                .sheet(Employee.class, Arrays.asList(employee), "People")
                .toWorkbook();
        try {
            final XSSFSheet sheet = (XSSFSheet) workbook.getSheet("People");
            assertThat(sheet.getDataValidations()).as("enum dropdown (data validation) should be created").isNotEmpty();
        } finally {
            workbook.close();
        }
    }

    // ------------------------------------------------------------------
    // import column-range bounds: exclude a duplicate header outside the range
    // ------------------------------------------------------------------

    @Test
    public void importColumnRange_unbounded_duplicateHeaderThrows() throws Exception {
        final byte[] bytes = buildDuplicateHeaderFixture();

        assertThrows(PxlDataException.class, () -> pxl.importExcel()
                .sheet(Employee.class, Arrays.asList("Dup"))
                .fromStream(new ByteArrayInputStream(bytes)));
    }

    @Test
    public void importColumnRange_bounded_excludesDuplicate() throws Exception {
        final byte[] bytes = buildDuplicateHeaderFixture();

        // limit data columns to 1..2 (1-based, index 0..1) -> exclude the duplicate "Name" at index 2
        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .importFirstDataColumnIndex(1)
                .importLastDataColumnIndex(2)
                .build();
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder()
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();

        final List<Employee> employees = pxl.importExcel()
                .override(option)
                .sheet(Employee.class, Arrays.asList("Dup"))
                .fromStream(new ByteArrayInputStream(bytes));

        assertThat(employees).hasSize(1);
        assertThat(employees.get(0).getName()).isEqualTo("Alice");
        assertThat(employees.get(0).getAge()).isEqualTo(30);
    }

    // ------------------------------------------------------------------
    // Workbook option: sheet-option list accessors (instance/static get, getList, add)
    // ------------------------------------------------------------------

    @Test
    public void importWorkbookOption_sheetOptionAccessors() throws PxlNullPointerException {
        final PxlImportSheetOption s0 = PxlImportSheetOption.builder().build();
        final PxlImportSheetOption s1 = PxlImportSheetOption.builder().build();
        final PxlImportWorkbookOption option = PxlImportWorkbookOption.builder().build();

        assertThat(option.addImportSheetOption(s0)).isTrue();
        option.addImportSheetOption(s1);

        // instance getter by index (out of range -> null)
        assertThat(option.getImportSheetOption(0)).isSameAs(s0);
        assertThat(option.getImportSheetOption(5)).isNull();

        // static null-safe accessors
        assertThat(PxlImportWorkbookOption.getImportSheetOption(option, 1)).isSameAs(s1);
        assertThat(PxlImportWorkbookOption.getImportSheetOption(null, 0)).isNull();
        assertThat(PxlImportWorkbookOption.getImportSheetOptions(option)).containsExactly(s0, s1);
        assertThat(PxlImportWorkbookOption.getImportSheetOptions(null)).isEmpty();

        assertThrows(PxlNullPointerException.class, () -> option.addImportSheetOption(null));
    }

    @Test
    public void exportWorkbookOption_sheetOptionAccessors() throws PxlNullPointerException {
        final PxlExportSheetOption s0 = PxlExportSheetOption.builder().build();
        final PxlExportSheetOption s1 = PxlExportSheetOption.builder().build();
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder().build();

        assertThat(option.addExportSheetOption(s0)).isTrue();
        option.addExportSheetOption(s1);

        assertThat(option.getExportSheetOption(0)).isSameAs(s0);
        assertThat(option.getExportSheetOption(5)).isNull();

        assertThat(PxlExportWorkbookOption.getExportSheetOption(option, 1)).isSameAs(s1);
        assertThat(PxlExportWorkbookOption.getExportSheetOption(null, 0)).isNull();
        assertThat(PxlExportWorkbookOption.getExportSheetOptions(option)).containsExactly(s0, s1);
        assertThat(PxlExportWorkbookOption.getExportSheetOptions(null)).isEmpty();

        assertThrows(PxlNullPointerException.class, () -> option.addExportSheetOption(null));
    }
}
