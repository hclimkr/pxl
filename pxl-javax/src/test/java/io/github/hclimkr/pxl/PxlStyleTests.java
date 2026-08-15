package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import io.github.hclimkr.pxl.styler.data.PxlDataHorizontalCenterTextStyler;
import io.github.hclimkr.pxl.styler.header.PxlHeaderHorizontalCenterTextStyler;
import io.github.hclimkr.pxl.styler.header.PxlHeaderWrapTextStyler;
import io.github.hclimkr.pxl.tcdata.*;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.List;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static org.assertj.core.api.Assertions.assertThat;

/**
 * Cell style / column width tests - header fill and font color, data alignment/wrap/border, fixed column width,
 * and workbook/sheet/column-level header and data styler cascade.
 */
public class PxlStyleTests {

    private static Pxl pxl;

    private static Workbook workbook;
    private static Sheet sheet;

    @BeforeAll
    public static void setUpBeforeClass() throws Exception {
        pxl = new Pxl();

        final StyledRow row = new StyledRow();
        row.setRequired("r");
        row.setOptional("o");
        row.setCentered("c");
        row.setWrapped("w");
        row.setBordered("b");
        row.setWide("wd");

        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder().exportDataValidation(false).build();
        workbook = pxl.exportExcel()
                .sheet(StyledRow.class, Arrays.asList(row), "Styled")
                .override(option)
                .toWorkbook();
        sheet = workbook.getSheet("Styled");
    }

    @AfterAll
    public static void tearDownAfterClass() throws Exception {
        if (workbook != null) {
            workbook.close();
        }
    }

    // header name -> column index
    private static int col(final String header) {
        final Row headerRow = sheet.getRow(0);
        for (final Cell cell : headerRow) {
            if (header.equals(cell.getStringCellValue())) {
                return cell.getColumnIndex();
            }
        }
        throw new IllegalStateException("header not found: " + header);
    }

    private static CellStyle headerStyle(final String header) {
        return sheet.getRow(0).getCell(col(header)).getCellStyle();
    }

    private static CellStyle dataStyle(final String header) {
        return sheet.getRow(1).getCell(col(header)).getCellStyle();
    }

    private static short fontColor(final CellStyle style) {
        final Font font = workbook.getFontAt(style.getFontIndexAsInt());
        return font.getColor();
    }

    private static int colIndex(final Sheet sheet, final String header) {
        for (final Cell cell : sheet.getRow(0)) {
            if (header.equals(cell.getStringCellValue())) {
                return cell.getColumnIndex();
            }
        }
        throw new IllegalStateException("header not found: " + header);
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

    @Test
    public void headerStyle_fill_isGreySolid() {
        final CellStyle style = headerStyle("Optional");
        assertThat(style.getFillPattern()).isEqualTo(PxlConstants.HEADER_COLUMN_FILL_PATTERN);
        assertThat(style.getFillForegroundColor()).isEqualTo(PxlConstants.HEADER_COLUMN_FOREGROUND_COLOR.getIndex());
    }

    @Test
    public void headerStyle_fontColor_requiredVsOptionalDiffer() {
        final short required = fontColor(headerStyle("Required"));   // @NotNull -> required
        final short optional = fontColor(headerStyle("Optional"));

        assertThat(required).isEqualTo(PxlConstants.REQUIRED_HEADER_COLUMN_FONT_COLOR.getIndex());
        assertThat(optional).isEqualTo(PxlConstants.OPTIONAL_HEADER_COLUMN_FONT_COLOR.getIndex());
        assertThat(required).isNotEqualTo(optional);
    }

    // ------------------------------------------------------------------
    // Data style: default (vertical center) + custom (horizontal center / wrap / border)
    // ------------------------------------------------------------------

    @Test
    public void dataStyle_default_verticalCenter() {
        // A column without a custom styler gets the default data styler (vertical center)
        assertThat(dataStyle("Required").getVerticalAlignment()).isEqualTo(VerticalAlignment.CENTER);
    }

    @Test
    public void dataStyle_custom_horizontalCenter() {
        assertThat(dataStyle("Centered").getAlignment()).isEqualTo(HorizontalAlignment.CENTER);
    }

    @Test
    public void dataStyle_custom_wrapText() {
        assertThat(dataStyle("Wrapped").getWrapText()).isTrue();
    }

    @Test
    public void dataStyle_custom_thinBorder() {
        final CellStyle style = dataStyle("Bordered");
        assertThat(style.getBorderTop()).isEqualTo(BorderStyle.THIN);
        assertThat(style.getBorderBottom()).isEqualTo(BorderStyle.THIN);
        assertThat(style.getBorderLeft()).isEqualTo(BorderStyle.THIN);
        assertThat(style.getBorderRight()).isEqualTo(BorderStyle.THIN);
    }

    // ------------------------------------------------------------------
    // Fixed column width
    // ------------------------------------------------------------------

    @Test
    public void columnWidth_fixed_applied() {
        assertThat(sheet.getColumnWidth(col("Wide"))).isEqualTo(5000);
    }

    // ------------------------------------------------------------------
    // Workbook-level header styler (option)
    // ------------------------------------------------------------------

    @Test
    public void headerStyler_workbookLevel_requiredAndOptional() throws Exception {
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportWorkbookRequiredHeaderCellStyler(PxlHeaderHorizontalCenterTextStyler.class)
                .exportWorkbookOptionalHeaderCellStyler(PxlHeaderWrapTextStyler.class)
                .build();

        final Workbook workbook = pxl.exportExcel()
                .sheet(HeaderStyleRow.class, Arrays.asList(sampleRow()), "T")
                .override(option)
                .toWorkbook();
        try {
            assertHeaderStyles(workbook, "T");
        } finally {
            workbook.close();
        }
    }

    // ------------------------------------------------------------------
    // Workbook-level stylers declared on @PxlWorkbook (the option path is covered above)
    // ------------------------------------------------------------------

    @Test
    public void stylers_workbookLevelAnnotation_applyToHeaderAndData() throws Exception {
        final WorkbookStylerWorkbook workbook = new WorkbookStylerWorkbook();
        workbook.setWorkbookName("W");
        workbook.setRows(Arrays.asList(sampleRow()));

        final Workbook poi = pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption())
                .toWorkbook();
        try {
            assertHeaderStyles(poi, "Styled");

            final Sheet styled = poi.getSheet("Styled");
            assertThat(styled.getRow(1).getCell(colIndex(styled, "Req")).getCellStyle().getAlignment())
                    .as("the workbook data styler reaches the data cells")
                    .isEqualTo(HorizontalAlignment.CENTER);
        } finally {
            poi.close();
        }
    }

    // ------------------------------------------------------------------
    // Sheet-level header styler (@PxlSheet)
    // ------------------------------------------------------------------

    @Test
    public void headerStyler_sheetLevel_requiredAndOptional() throws Exception {
        final SheetHeaderStylerWorkbook workbook = new SheetHeaderStylerWorkbook();
        workbook.setWorkbookName("W");
        workbook.setRows(Arrays.asList(sampleRow()));

        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder().exportDataValidation(false).build();
        final Workbook poi = pxl.exportExcel()
                .workbook(workbook)
                .override(option)
                .toWorkbook();
        try {
            assertHeaderStyles(poi, "Header");
        } finally {
            poi.close();
        }
    }

    // ------------------------------------------------------------------
    // Sheet data cell styler cascade (exportSheetDataCellStyler)
    // ------------------------------------------------------------------

    @Test
    public void dataStyler_sheetLevel_appliesToColumns() throws Exception {
        final SheetStylerWorkbook workbook = new SheetStylerWorkbook();
        workbook.setWorkbookName("W");
        workbook.setRows(twoEmployees());

        final Workbook poi = pxl.exportExcel()
                .workbook(workbook)
                .override(noValidationOption())
                .toWorkbook();
        try {
            final Sheet sheet = poi.getSheet("Centered");
            // The data cell inherits the sheet styler (horizontal center)
            assertThat(sheet.getRow(1).getCell(colIndex(sheet, "Name")).getCellStyle().getAlignment())
                    .isEqualTo(HorizontalAlignment.CENTER);
        } finally {
            poi.close();
        }
    }

    // ------------------------------------------------------------------
    // Workbook data cell styler cascade (option)
    // ------------------------------------------------------------------

    @Test
    public void dataStyler_workbookLevel_appliesViaOption() throws Exception {
        final PxlExportWorkbookOption option = PxlExportWorkbookOption.builder()
                .exportDataValidation(false)
                .exportWorkbookDataCellStyler(PxlDataHorizontalCenterTextStyler.class)
                .build();

        final Workbook poi = pxl.exportExcel()
                .sheet(Employee.class, twoEmployees(), "People")
                .override(option)
                .toWorkbook();
        try {
            final Sheet sheet = poi.getSheet("People");
            assertThat(sheet.getRow(1).getCell(colIndex(sheet, "Name")).getCellStyle().getAlignment())
                    .isEqualTo(HorizontalAlignment.CENTER);
        } finally {
            poi.close();
        }
    }

    // ------------------------------------------------------------------
    // Column header styler (exportColumnRequiredHeaderCellStyler)
    // ------------------------------------------------------------------

    @Test
    public void headerStyler_columnLevel_applied() throws Exception {
        final ColumnHeaderStylerRow row = new ColumnHeaderStylerRow();
        row.setCustom("x");

        final Workbook poi = pxl.exportExcel()
                .sheet(ColumnHeaderStylerRow.class, Arrays.asList(row), "T")
                .override(noValidationOption())
                .toWorkbook();
        try {
            final Sheet sheet = poi.getSheet("T");
            // The header cell applies the custom required header styler (horizontal center)
            assertThat(sheet.getRow(0).getCell(colIndex(sheet, "Custom")).getCellStyle().getAlignment())
                    .isEqualTo(HorizontalAlignment.CENTER);
        } finally {
            poi.close();
        }
    }

    // ------------------------------------------------------------------
    // Column header styler (exportColumnOptionalHeaderCellStyler)
    // ------------------------------------------------------------------

    @Test
    public void headerStyler_columnLevelOptional_applied() throws Exception {
        // Which of the two column stylers is consulted follows from the column's own constraints, so the optional
        // one needs a column without them: Plain carries no @NotNull and must pick up the wrap-text styler while
        // the @NotNull column beside it keeps the required one.
        final ColumnHeaderStylerRow row = new ColumnHeaderStylerRow();
        row.setCustom("x");
        row.setPlain("y");

        final Workbook poi = pxl.exportExcel()
                .sheet(ColumnHeaderStylerRow.class, Arrays.asList(row), "T")
                .override(noValidationOption())
                .toWorkbook();
        try {
            final Sheet sheet = poi.getSheet("T");
            assertThat(sheet.getRow(0).getCell(colIndex(sheet, "Plain")).getCellStyle().getWrapText())
                    .as("the optional header styler applies to a column without a required constraint")
                    .isTrue();
            assertThat(sheet.getRow(0).getCell(colIndex(sheet, "Custom")).getCellStyle().getWrapText())
                    .as("the required column keeps its own styler")
                    .isFalse();
        } finally {
            poi.close();
        }
    }
}
