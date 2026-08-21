package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.exception.PxlValidationException;
import io.github.hclimkr.pxl.option.PxlExportColumnOption;
import io.github.hclimkr.pxl.option.PxlExportSheetOption;
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import io.github.hclimkr.pxl.option.PxlImportColumnOption;
import io.github.hclimkr.pxl.tcdata.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
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
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static io.github.hclimkr.pxl.tcdata.TestExports.emit;
import static io.github.hclimkr.pxl.tcdata.TestExports.workbookOf;
import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Tests for various @PxlColumn property combinations.
 * <p>
 * Verifies pattern/importPattern/exportPattern, custom true/false strings, custom collection separators, exportNullString,
 * importTrim, column name aliases (name={...}), importUnique, exportOptionItems, exportEnumDropDownListStyle,
 * exportSampleEnabled against a data export, and column-level import/export enable toggles via actual
 * round-trips/assertions.
 * <p>
 * Every test that exports is swept across {@link ExportDest}: what a column option renders has to be the same on
 * every terminal. The import-only tests build their input with raw POI and never export at all.
 */
public class PxlColumnOptionTests {

    private static Pxl pxl;

    private static final DataFormatter DATA_FORMATTER = new DataFormatter();

    @BeforeAll
    public static void setUpBeforeClass() {
        pxl = new Pxl();
    }

    private TestInfo testInfo;

    @BeforeEach
    public void bindTestInfo(final TestInfo testInfo) {
        this.testInfo = testInfo;
    }

    private <T> List<T> roundTrip(final ExportDest dest, final String sheetName, final List<T> rows, final Class<T> rowClass) throws Exception {
        final byte[] bytes = exportBytes(dest, sheetName, rows, rowClass, noValidationOption());
        return pxl.importExcel()
                .sheet(rowClass, Arrays.asList(sheetName))
                .fromStream(new ByteArrayInputStream(bytes));
    }

    private <T> byte[] exportBytes(final ExportDest dest, final String sheetName, final List<T> rows, final Class<T> rowClass,
                                   final PxlExportWorkbookOption option) throws Exception {
        return emit(pxl.exportExcel()
                .sheet(rowClass, rows, sheetName)
                .override(option), dest, testInfo);
    }

    // Builds a STRING sheet with 1 header row + N data rows via POI.
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

    // The list of headers in the given sheet's header row
    private static List<String> headersOf(final byte[] bytes, final String sheetName) throws Exception {
        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            final Row header = workbook.getSheet(sheetName).getRow(0);
            final List<String> headers = new ArrayList<>();
            for (final Cell cell : header) {
                headers.add(cell.getStringCellValue());
            }
            return headers;
        }
    }

    // The rendered value of a data cell in the given sheet by (row, header)
    private static String renderedCell(final byte[] bytes, final String sheetName, final int dataRowIndex, final String header) throws Exception {
        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(bytes))) {
            final Sheet sheet = workbook.getSheet(sheetName);
            final Row headerRow = sheet.getRow(0);
            int col = -1;
            for (final Cell cell : headerRow) {
                if (header.equals(cell.getStringCellValue())) {
                    col = cell.getColumnIndex();
                    break;
                }
            }
            assertThat(col).as("header '" + header + "' not found").isGreaterThanOrEqualTo(0);
            final Cell cell = sheet.getRow(dataRowIndex).getCell(col);
            return cell == null ? "" : DATA_FORMATTER.formatCellValue(cell);
        }
    }

    // ------------------------------------------------------------------
    // pattern / importPattern / exportPattern
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void pattern_customFormat_roundTrips(final ExportDest dest) throws Exception {
        final PatternRow row = new PatternRow();
        row.setDate(LocalDate.of(2023, 6, 15));
        row.setTime(LocalTime.of(10, 30, 45));
        row.setAmount(new BigDecimal("1234.5"));
        row.setTimestamp(LocalDateTime.of(2023, 6, 15, 10, 30));   // minute granularity (no seconds in the pattern)

        // verify the exported cell rendering follows the custom pattern
        final byte[] bytes = exportBytes(dest, "Pattern", Arrays.asList(row), PatternRow.class, noValidationOption());
        assertThat(renderedCell(bytes, "Pattern", 1, "Date")).isEqualTo("2023/06/15");
        assertThat(renderedCell(bytes, "Pattern", 1, "Time")).isEqualTo("10.30.45");
        assertThat(renderedCell(bytes, "Pattern", 1, "Amount")).isEqualTo("1,234.50");
        assertThat(renderedCell(bytes, "Pattern", 1, "Timestamp")).isEqualTo("2023.06.15 10:30");

        // round-trip value preservation
        final PatternRow out = roundTrip(dest, "Pattern", Arrays.asList(row), PatternRow.class).get(0);
        assertThat(out.getDate()).isEqualTo(LocalDate.of(2023, 6, 15));
        assertThat(out.getTime()).isEqualTo(LocalTime.of(10, 30, 45));
        assertThat(out.getAmount()).isEqualByComparingTo("1234.5");
        assertThat(out.getTimestamp()).isEqualTo(LocalDateTime.of(2023, 6, 15, 10, 30));
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void pattern_numberSymbols_localeIndependent_useRootSymbols(final ExportDest dest) throws Exception {
        // A numeric pattern's decimal/grouping symbols must not depend on the JVM default locale. Under a
        // comma-decimal locale (de_DE, whose default symbols are ',' decimal / '.' grouping) the "#,##0.00"
        // pattern must still render with Locale.ROOT symbols ('.' decimal / ',' grouping), and round-trip.
        final Locale previous = Locale.getDefault();
        try {
            Locale.setDefault(Locale.GERMANY);

            final PatternRow row = new PatternRow();
            row.setDate(LocalDate.of(2023, 6, 15));
            row.setTime(LocalTime.of(10, 30, 45));
            row.setAmount(new BigDecimal("1234.5"));
            row.setTimestamp(LocalDateTime.of(2023, 6, 15, 10, 30));

            final byte[] bytes = exportBytes(dest, "Pattern", Arrays.asList(row), PatternRow.class, noValidationOption());
            // Locale.ROOT symbols regardless of the de_DE default (would otherwise be "1.234,50").
            assertThat(renderedCell(bytes, "Pattern", 1, "Amount")).isEqualTo("1,234.50");

            // Import DecimalFormat also uses Locale.ROOT, so the value round-trips under the de_DE default.
            final PatternRow out = roundTrip(dest, "Pattern", Arrays.asList(row), PatternRow.class).get(0);
            assertThat(out.getAmount()).isEqualByComparingTo("1234.5");
        } finally {
            Locale.setDefault(previous);
        }
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void dateTimeTypes_noPattern_export_writesLocaleIndependentIso(final ExportDest dest) throws Exception {
        // Without a column pattern, a date/time value's string form (Collection elements are always string cells)
        // is written as fixed ISO-8601, not the JVM default locale's format (e.g. ko_KR "2023. 1. 2."). The value
        // round-trips on any machine because the read patterns are fixed ISO too (no locale dependence).
        final CollectionTypesRow row = new CollectionTypesRow();
        row.setLocalDates(Arrays.asList(LocalDate.of(2023, 1, 2)));
        row.setLocalTimes(Arrays.asList(LocalTime.of(3, 4, 5)));
        row.setLocalDateTimes(Arrays.asList(LocalDateTime.of(2023, 1, 2, 3, 4, 5)));

        final byte[] bytes = exportBytes(dest, "C", Arrays.asList(row), CollectionTypesRow.class, noValidationOption());
        assertThat(renderedCell(bytes, "C", 1, "LocalDates")).isEqualTo("2023-01-02");
        assertThat(renderedCell(bytes, "C", 1, "LocalTimes")).isEqualTo("03:04:05");
        assertThat(renderedCell(bytes, "C", 1, "LocalDateTimes")).isEqualTo("2023-01-02T03:04:05");

        final CollectionTypesRow out = roundTrip(dest, "C", Arrays.asList(row), CollectionTypesRow.class).get(0);
        assertThat(out.getLocalDates()).containsExactly(LocalDate.of(2023, 1, 2));
        assertThat(out.getLocalTimes()).containsExactly(LocalTime.of(3, 4, 5));
        assertThat(out.getLocalDateTimes()).containsExactly(LocalDateTime.of(2023, 1, 2, 3, 4, 5));
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void pattern_dateTextField_localeIndependent_useRootMonthName(final ExportDest dest) throws Exception {
        // A date pattern's text fields (month/day names, AM/PM) are resolved with Locale.ROOT, so they do not
        // depend on the JVM default locale. Under de_DE the "yyyy-MMM-dd" month name is still the ROOT/English
        // "Jun" ("2023-Jun-15"), not the German "Juni".
        final Locale previous = Locale.getDefault();
        try {
            Locale.setDefault(Locale.GERMANY);

            final PatternRow row = new PatternRow();
            row.setDate(LocalDate.of(2023, 6, 15));

            // Override the "date" column's export pattern with one that has a text (month-name) field.
            final PxlExportColumnOption dateOption = PxlExportColumnOption.builder()
                    .fieldName("date")
                    .exportPattern("yyyy-MMM-dd")
                    .build();
            final PxlExportSheetOption sheetOption = PxlExportSheetOption.builder()
                    .exportColumnOptions(Arrays.asList(dateOption))
                    .build();
            final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                    .exportDataValidation(false)
                    .exportSheetOptions(Arrays.asList(sheetOption))
                    .build();

            final byte[] bytes = exportBytes(dest, "Pattern", Arrays.asList(row), PatternRow.class, option);
            assertThat(renderedCell(bytes, "Pattern", 1, "Date")).isEqualTo("2023-Jun-15");
        } finally {
            Locale.setDefault(previous);
        }
    }

    // ------------------------------------------------------------------
    // Custom true/false strings
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void booleanStrings_customYN_roundTrips(final ExportDest dest) throws Exception {
        final BoolStringRow t = new BoolStringRow();
        t.setFlag(Boolean.TRUE);
        final BoolStringRow f = new BoolStringRow();
        f.setFlag(Boolean.FALSE);

        // exported cells are rendered as Y/N
        final byte[] bytes = exportBytes(dest, "Bool", Arrays.asList(t, f), BoolStringRow.class, noValidationOption());
        assertThat(renderedCell(bytes, "Bool", 1, "Flag")).isEqualTo("Y");
        assertThat(renderedCell(bytes, "Bool", 2, "Flag")).isEqualTo("N");

        // round-trip
        final List<BoolStringRow> out = roundTrip(dest, "Bool", Arrays.asList(t, f), BoolStringRow.class);
        assertThat(out.get(0).getFlag()).isTrue();
        assertThat(out.get(1).getFlag()).isFalse();
    }

    // ------------------------------------------------------------------
    // Custom collection separator
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void collectionSeparator_custom_roundTrips(final ExportDest dest) throws Exception {
        final SeparatorRow row = new SeparatorRow();
        row.setTags(Arrays.asList("red", "green", "blue"));
        row.setNums(Arrays.asList(1, 2, 3));

        final byte[] bytes = exportBytes(dest, "Sep", Arrays.asList(row), SeparatorRow.class, noValidationOption());
        assertThat(renderedCell(bytes, "Sep", 1, "Tags")).isEqualTo("red|green|blue");
        assertThat(renderedCell(bytes, "Sep", 1, "Nums")).isEqualTo("1/2/3");

        final SeparatorRow out = roundTrip(dest, "Sep", Arrays.asList(row), SeparatorRow.class).get(0);
        assertThat(out.getTags()).containsExactly("red", "green", "blue");
        assertThat(out.getNums()).containsExactly(1, 2, 3);
    }

    // ------------------------------------------------------------------
    // exportNullString
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportNullString_custom_rendered(final ExportDest dest) throws Exception {
        final NullStringRow row = new NullStringRow();
        row.setValue(null);
        row.setLabel("keep");

        final byte[] bytes = exportBytes(dest, "Null", Arrays.asList(row), NullStringRow.class, noValidationOption());
        assertThat(renderedCell(bytes, "Null", 1, "Value")).isEqualTo("N/A");
        assertThat(renderedCell(bytes, "Null", 1, "Label")).isEqualTo("keep");
    }

    // ------------------------------------------------------------------
    // importTrim (false = preserve whitespace)
    // ------------------------------------------------------------------

    @Test
    public void importTrim_disabled_preservesWhitespace() throws Exception {
        final byte[] bytes = buildStringSheet("Trim",
                new String[]{"Raw", "Trimmed"},
                new String[][]{{"  spaced  ", "  spaced  "}});

        final List<ImportTrimRow> rows = pxl.importExcel()
                .sheet(ImportTrimRow.class, Arrays.asList("Trim"))
                .fromStream(new ByteArrayInputStream(bytes));

        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).getRaw()).isEqualTo("  spaced  ");   // not trimmed
        assertThat(rows.get(0).getTrimmed()).isEqualTo("spaced");   // default trim
    }

    // ------------------------------------------------------------------
    // importUnique
    // ------------------------------------------------------------------

    @Test
    public void importUnique_duplicateValue_throws() throws Exception {
        final byte[] bytes = buildStringSheet("Unique",
                new String[]{"Code", "Name"},
                new String[][]{{"C1", "a"}, {"C1", "b"}});   // duplicate Code

        assertThrows(PxlValidationException.class, () -> pxl.importExcel()
                .sheet(UniqueCodeRow.class, Arrays.asList("Unique"))
                .fromStream(new ByteArrayInputStream(bytes)));
    }

    @Test
    public void importUnique_distinctValues_ok() throws Exception {
        final byte[] bytes = buildStringSheet("Unique",
                new String[]{"Code", "Name"},
                new String[][]{{"C1", "a"}, {"C2", "b"}});

        final List<UniqueCodeRow> rows = pxl.importExcel()
                .sheet(UniqueCodeRow.class, Arrays.asList("Unique"))
                .fromStream(new ByteArrayInputStream(bytes));

        assertThat(rows).hasSize(2);
        assertThat(rows).extracting(UniqueCodeRow::getCode).containsExactly("C1", "C2");
    }

    // ------------------------------------------------------------------
    // exportOptionItems (fixed-list dropdown)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void exportOptionItems_dropdownPresent(final ExportDest dest) throws Exception {
        final OptionItemsRow row = new OptionItemsRow();
        row.setChoice("Red");

        // option null -> exportDataValidation defaults to true -> dropdown created
        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(OptionItemsRow.class, Arrays.asList(row), "Opt"), dest, testInfo)) {
            final XSSFSheet sheet = (XSSFSheet) workbook.getSheet("Opt");
            assertThat(sheet.getDataValidations()).as("exportOptionItems dropdown should be created").isNotEmpty();
        }
    }

    // ------------------------------------------------------------------
    // exportEnumDropDownListStyle (NONE suppresses, SORTED_SET creates)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void enumDropdownStyle_variants_applied(final ExportDest dest) throws Exception {
        final EnumStyleRow row = new EnumStyleRow();
        row.setGradeNone(Grade.A);
        row.setGradeSorted(Grade.B);

        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(EnumStyleRow.class, Arrays.asList(row), "EnumStyle"), dest, testInfo)) {
            final XSSFSheet sheet = (XSSFSheet) workbook.getSheet("EnumStyle");
            // the NONE column has no dropdown, only the SORTED_SET column does -> 1 data validation
            assertThat(sheet.getDataValidations()).hasSize(1);
        }
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void enumDropdownWithOptionItems_setKeepsOrder_sortedSetSorts(final ExportDest dest) throws Exception {
        // An enum column with an explicit exportOptionItems list: SET uses the items as given,
        // SORTED_SET sorts them.
        final EnumOptionItemsRow row = new EnumOptionItemsRow();
        row.setGradeSet(Grade.A);
        row.setGradeSorted(Grade.B);

        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(EnumOptionItemsRow.class, Arrays.asList(row), "Dropdowns"), dest, testInfo)) {
            final XSSFSheet sheet = (XSSFSheet) workbook.getSheet("Dropdowns");
            assertThat(sheet.getDataValidations()).hasSize(2);

            final List<List<String>> optionLists = new ArrayList<>();
            for (final XSSFDataValidation dataValidation : sheet.getDataValidations()) {
                optionLists.add(Arrays.asList(dataValidation.getValidationConstraint().getExplicitListValues()));
            }
            // SORTED_SET sorted {"C","A","B"} -> [A,B,C]; SET kept {"B","A","C"} -> [B,A,C]
            assertThat(optionLists).contains(Arrays.asList("A", "B", "C"), Arrays.asList("B", "A", "C"));
        }
    }

    // ------------------------------------------------------------------
    // Column-level import/export enable toggles
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void columnToggle_exportAndImportOff_excluded(final ExportDest dest) throws Exception {
        final ColumnToggleRow row = new ColumnToggleRow();
        row.setAlways("a");
        row.setExportOff("b");
        row.setImportOff("c");

        final byte[] bytes = exportBytes(dest, "Toggle", Arrays.asList(row), ColumnToggleRow.class, noValidationOption());

        // an exportEnabled=false column is not exported.
        final List<String> headers = headersOf(bytes, "Toggle");
        assertThat(headers).contains("Always", "ImportOff");
        assertThat(headers).doesNotContain("ExportOff");

        final List<ColumnToggleRow> rows = pxl.importExcel()
                .sheet(ColumnToggleRow.class, Arrays.asList("Toggle"))
                .fromStream(new ByteArrayInputStream(bytes));

        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).getAlways()).isEqualTo("a");
        // an importEnabled=false column is not imported even when the header is present.
        assertThat(rows.get(0).getImportOff()).isNull();
        // no value either, since it was not exported
        assertThat(rows.get(0).getExportOff()).isNull();
    }

    // ------------------------------------------------------------------
    // exportSampleEnabled governs the sample export alone
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void columnSampleDisabled_dataExport_stillWritesColumn(final ExportDest dest) throws Exception {
        // The sample export drops the Skip column (covered in PxlSampleExcelExportTests); a data export must not,
        // or the two flags would collapse into one.
        final SampleColumnRow row = new SampleColumnRow();
        row.setKeep("k");
        row.setSkip("s");

        final SampleColumnRow imported = roundTrip(dest, "Data", Arrays.asList(row), SampleColumnRow.class).get(0);

        assertThat(imported.getKeep()).isEqualTo("k");
        assertThat(imported.getSkip()).as("exportSampleEnabled=false is not exportEnabled=false").isEqualTo("s");
    }

    // ------------------------------------------------------------------
    // Column option: column-name normalization accessors
    // ------------------------------------------------------------------

    @Test
    public void importColumnOption_getImportColumnNames_stripsAllWhitespaceAndBlanks() {
        // Import normalization deletes ALL whitespace (inner too) and drops blank entries.
        final PxlImportColumnOption option = PxlImportColumnOption.builder()
                .fieldName("f").importColumnNames(Arrays.asList(" A B ", "", "c")).build();
        assertThat(PxlImportColumnOption.getImportColumnNames(option)).containsExactly("AB", "c");
        assertThat(PxlImportColumnOption.getImportColumnNames(null)).isNull();
        assertThat(PxlImportColumnOption.getImportColumnNames(
                PxlImportColumnOption.builder().fieldName("f").build())).isNull();   // names == null -> null
    }

    @Test
    public void exportColumnOption_getExportColumnNames_trimsAndDropsBlanks() {
        // Export normalization only trims the ends (inner whitespace kept) and drops blank entries.
        final PxlExportColumnOption option = PxlExportColumnOption.builder()
                .fieldName("f").exportColumnNames(Arrays.asList(" A B ", "", "c ")).build();
        assertThat(PxlExportColumnOption.getExportColumnNames(option)).containsExactly("A B", "c");
        assertThat(PxlExportColumnOption.getExportColumnNames(null)).isNull();
        assertThat(PxlExportColumnOption.getExportColumnNames(
                PxlExportColumnOption.builder().fieldName("f").build())).isNull();
    }
}
