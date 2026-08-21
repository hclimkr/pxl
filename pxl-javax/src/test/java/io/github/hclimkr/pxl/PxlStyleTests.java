package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import io.github.hclimkr.pxl.styler.data.PxlDataHorizontalCenterTextStyler;
import io.github.hclimkr.pxl.styler.header.PxlHeaderHorizontalCenterTextStyler;
import io.github.hclimkr.pxl.styler.header.PxlHeaderWrapTextStyler;
import io.github.hclimkr.pxl.tcdata.*;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.TestInfo;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;

import java.util.Arrays;
import java.util.List;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static io.github.hclimkr.pxl.tcdata.TestExports.workbookOf;
import static org.assertj.core.api.Assertions.assertThat;

/**
 * Cell style / column width tests - header fill and font color, data alignment/wrap/border, fixed column width,
 * and workbook/sheet/column-level header and data styler cascade.
 * <p>
 * A style has to survive whichever terminal produced the workbook, so every test here is swept across
 * {@link ExportDest}. That is also why the styled workbook is built per test rather than once in a
 * {@code @BeforeAll}: the destination is a parameter now, so there is no single fixture to share.
 */
public class PxlStyleTests {

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

    // The workbook the style assertions in the first half read, exported to the given destination.
    private Workbook styledWorkbook(final ExportDest dest) throws Exception {
        final StyledRow row = new StyledRow();
        row.setRequired("r");
        row.setOptional("o");
        row.setCentered("c");
        row.setWrapped("w");
        row.setBordered("b");
        row.setWide("wd");

        return workbookOf(pxl.exportExcel()
                .sheet(StyledRow.class, Arrays.asList(row), "Styled")
                .override(PxlExportWorkbookOption.builder().exportDataValidation(false).build()), dest, testInfo);
    }

    // header name -> column index
    private static int colIndex(final Sheet sheet, final String header) {
        for (final Cell cell : sheet.getRow(0)) {
            if (header.equals(cell.getStringCellValue())) {
                return cell.getColumnIndex();
            }
        }
        throw new IllegalStateException("header not found: " + header);
    }

    private static CellStyle headerStyle(final Sheet sheet, final String header) {
        return sheet.getRow(0).getCell(colIndex(sheet, header)).getCellStyle();
    }

    private static CellStyle dataStyle(final Sheet sheet, final String header) {
        return sheet.getRow(1).getCell(colIndex(sheet, header)).getCellStyle();
    }

    private static short fontColor(final Workbook workbook, final CellStyle style) {
        final Font font = workbook.getFontAt(style.getFontIndexAsInt());
        return font.getColor();
    }

    private static HeaderStyleRow sampleRow() {
        final HeaderStyleRow row = new HeaderStyleRow();
        row.setReq("r");
        row.setOpt("o");
        return row;
    }

    // Verifies the required/optional styler application via the header cell's horizontal alignment / wrap.
    private static void assertHeaderStyles(final Workbook workbook, final String sheetName) {
        final Sheet sheet = workbook.getSheet(sheetName);
        assertThat(sheet).isNotNull();

        // Required (Req) header -> horizontal center
        assertThat(sheet.getRow(0).getCell(colIndex(sheet, "Req")).getCellStyle().getAlignment())
                .as("required header styler (horizontal center)").isEqualTo(HorizontalAlignment.CENTER);
        // Optional (Opt) header -> wrap text
        assertThat(sheet.getRow(0).getCell(colIndex(sheet, "Opt")).getCellStyle().getWrapText())
                .as("optional header styler (wrap text)").isTrue();
    }

    private static List<Employee> twoEmployees() {
        return Arrays.asList(
                Fixtures.employee("Alice", 30, "50000", true, null, Grade.A, "Engineering"),
                Fixtures.employee("Bob", 42, "72000", false, null, Grade.B, "Sales"));
    }

    // ------------------------------------------------------------------
    // Header style: fill background and required/optional font color
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void headerStyle_fill_isGreySolid(final ExportDest dest) throws Exception {
        try (Workbook workbook = styledWorkbook(dest)) {
            final CellStyle style = headerStyle(workbook.getSheet("Styled"), "Optional");
            assertThat(style.getFillPattern()).isEqualTo(PxlConstants.HEADER_COLUMN_FILL_PATTERN);
            assertThat(style.getFillForegroundColor()).isEqualTo(PxlConstants.HEADER_COLUMN_FOREGROUND_COLOR.getIndex());
        }
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void headerStyle_fontColor_requiredVsOptionalDiffer(final ExportDest dest) throws Exception {
        try (Workbook workbook = styledWorkbook(dest)) {
            final Sheet sheet = workbook.getSheet("Styled");
            final short required = fontColor(workbook, headerStyle(sheet, "Required"));   // @NotNull -> required
            final short optional = fontColor(workbook, headerStyle(sheet, "Optional"));

            assertThat(required).isEqualTo(PxlConstants.REQUIRED_HEADER_COLUMN_FONT_COLOR.getIndex());
            assertThat(optional).isEqualTo(PxlConstants.OPTIONAL_HEADER_COLUMN_FONT_COLOR.getIndex());
            assertThat(required).isNotEqualTo(optional);
        }
    }

    // ------------------------------------------------------------------
    // Data style: default (vertical center) + custom (horizontal center / wrap / border)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void dataStyle_default_verticalCenter(final ExportDest dest) throws Exception {
        // A column without a custom styler gets the default data styler (vertical center)
        try (Workbook workbook = styledWorkbook(dest)) {
            assertThat(dataStyle(workbook.getSheet("Styled"), "Required").getVerticalAlignment())
                    .isEqualTo(VerticalAlignment.CENTER);
        }
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void dataStyle_custom_horizontalCenter(final ExportDest dest) throws Exception {
        try (Workbook workbook = styledWorkbook(dest)) {
            assertThat(dataStyle(workbook.getSheet("Styled"), "Centered").getAlignment())
                    .isEqualTo(HorizontalAlignment.CENTER);
        }
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void dataStyle_custom_wrapText(final ExportDest dest) throws Exception {
        try (Workbook workbook = styledWorkbook(dest)) {
            assertThat(dataStyle(workbook.getSheet("Styled"), "Wrapped").getWrapText()).isTrue();
        }
    }

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void dataStyle_custom_thinBorder(final ExportDest dest) throws Exception {
        try (Workbook workbook = styledWorkbook(dest)) {
            final CellStyle style = dataStyle(workbook.getSheet("Styled"), "Bordered");
            assertThat(style.getBorderTop()).isEqualTo(BorderStyle.THIN);
            assertThat(style.getBorderBottom()).isEqualTo(BorderStyle.THIN);
            assertThat(style.getBorderLeft()).isEqualTo(BorderStyle.THIN);
            assertThat(style.getBorderRight()).isEqualTo(BorderStyle.THIN);
        }
    }

    // ------------------------------------------------------------------
    // Fixed column width
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void columnWidth_fixed_applied(final ExportDest dest) throws Exception {
        try (Workbook workbook = styledWorkbook(dest)) {
            final Sheet sheet = workbook.getSheet("Styled");
            assertThat(sheet.getColumnWidth(colIndex(sheet, "Wide"))).isEqualTo(5000);
        }
    }

    // ------------------------------------------------------------------
    // Workbook-level header styler (option)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void headerStyler_workbookLevel_requiredAndOptional(final ExportDest dest) throws Exception {
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportWorkbookRequiredHeaderCellStyler(PxlHeaderHorizontalCenterTextStyler.class)
                .exportWorkbookOptionalHeaderCellStyler(PxlHeaderWrapTextStyler.class)
                .build();

        try (Workbook workbook = workbookOf(pxl.exportExcel()
                .sheet(HeaderStyleRow.class, Arrays.asList(sampleRow()), "T")
                .override(option), dest, testInfo)) {
            assertHeaderStyles(workbook, "T");
        }
    }

    // ------------------------------------------------------------------
    // Workbook-level stylers declared on @PxlWorkbook (the option path is covered above)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void stylers_workbookLevelAnnotation_applyToHeaderAndData(final ExportDest dest) throws Exception {
        final WorkbookStylerWorkbook workbook = new WorkbookStylerWorkbook();
        workbook.setWorkbookName("W");
        workbook.setRows(Arrays.asList(sampleRow()));

        try (Workbook poi = workbookOf(pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption()), dest, testInfo)) {
            assertHeaderStyles(poi, "Styled");

            final Sheet styled = poi.getSheet("Styled");
            assertThat(styled.getRow(1).getCell(colIndex(styled, "Req")).getCellStyle().getAlignment())
                    .as("the workbook data styler reaches the data cells")
                    .isEqualTo(HorizontalAlignment.CENTER);
        }
    }

    // ------------------------------------------------------------------
    // Sheet-level header styler (@PxlSheet)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void headerStyler_sheetLevel_requiredAndOptional(final ExportDest dest) throws Exception {
        final SheetHeaderStylerWorkbook workbook = new SheetHeaderStylerWorkbook();
        workbook.setWorkbookName("W");
        workbook.setRows(Arrays.asList(sampleRow()));

        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder().exportDataValidation(false).build();
        try (Workbook poi = workbookOf(pxl.exportExcel()
                .workbook(workbook)
                .override(option), dest, testInfo)) {
            assertHeaderStyles(poi, "Header");
        }
    }

    // ------------------------------------------------------------------
    // Sheet data cell styler cascade (exportSheetDataCellStyler)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void dataStyler_sheetLevel_appliesToColumns(final ExportDest dest) throws Exception {
        final SheetStylerWorkbook workbook = new SheetStylerWorkbook();
        workbook.setWorkbookName("W");
        workbook.setRows(twoEmployees());

        try (Workbook poi = workbookOf(pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption()), dest, testInfo)) {
            final Sheet sheet = poi.getSheet("Centered");
            // The data cell inherits the sheet styler (horizontal center)
            assertThat(sheet.getRow(1).getCell(colIndex(sheet, "Name")).getCellStyle().getAlignment())
                    .isEqualTo(HorizontalAlignment.CENTER);
        }
    }

    // ------------------------------------------------------------------
    // Workbook data cell styler cascade (option)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void dataStyler_workbookLevel_appliesViaOption(final ExportDest dest) throws Exception {
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportWorkbookDataCellStyler(PxlDataHorizontalCenterTextStyler.class)
                .build();

        try (Workbook poi = workbookOf(pxl.exportExcel()
                .sheet(Employee.class, twoEmployees(), "People")
                .override(option), dest, testInfo)) {
            final Sheet sheet = poi.getSheet("People");
            assertThat(sheet.getRow(1).getCell(colIndex(sheet, "Name")).getCellStyle().getAlignment())
                    .isEqualTo(HorizontalAlignment.CENTER);
        }
    }

    // ------------------------------------------------------------------
    // Column header styler (exportColumnRequiredHeaderCellStyler)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void headerStyler_columnLevel_applied(final ExportDest dest) throws Exception {
        final ColumnHeaderStylerRow row = new ColumnHeaderStylerRow();
        row.setCustom("x");

        try (Workbook poi = workbookOf(pxl.exportExcel()
                .sheet(ColumnHeaderStylerRow.class, Arrays.asList(row), "T")
                .override(noValidationOption()), dest, testInfo)) {
            final Sheet sheet = poi.getSheet("T");
            // The header cell applies the custom required header styler (horizontal center)
            assertThat(sheet.getRow(0).getCell(colIndex(sheet, "Custom")).getCellStyle().getAlignment())
                    .isEqualTo(HorizontalAlignment.CENTER);
        }
    }

    // ------------------------------------------------------------------
    // Column header styler (exportColumnOptionalHeaderCellStyler)
    // ------------------------------------------------------------------

    @ParameterizedTest
    @EnumSource(ExportDest.class)
    public void headerStyler_columnLevelOptional_applied(final ExportDest dest) throws Exception {
        // Which of the two column stylers is consulted follows from the column's own constraints, so the optional
        // one needs a column without them: Plain carries no @NotNull and must pick up the wrap-text styler while
        // the @NotNull column beside it keeps the required one.
        final ColumnHeaderStylerRow row = new ColumnHeaderStylerRow();
        row.setCustom("x");
        row.setPlain("y");

        try (Workbook poi = workbookOf(pxl.exportExcel()
                .sheet(ColumnHeaderStylerRow.class, Arrays.asList(row), "T")
                .override(noValidationOption()), dest, testInfo)) {
            final Sheet sheet = poi.getSheet("T");
            assertThat(sheet.getRow(0).getCell(colIndex(sheet, "Plain")).getCellStyle().getWrapText())
                    .as("the optional header styler applies to a column without a required constraint")
                    .isTrue();
            assertThat(sheet.getRow(0).getCell(colIndex(sheet, "Custom")).getCellStyle().getWrapText())
                    .as("the required column keeps its own styler")
                    .isFalse();
        }
    }
}
