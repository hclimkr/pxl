package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import io.github.hclimkr.pxl.tcdata.PicturePrecedenceRow;
import io.github.hclimkr.pxl.tcdata.PictureRow;
import io.github.hclimkr.pxl.tcdata.TestPaths;
import io.github.hclimkr.pxl.type.PxlExcelEngine;
import io.github.hclimkr.pxl.util.PxlCellUtils;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInfo;
import org.junit.jupiter.api.io.TempDir;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.File;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static org.assertj.core.api.Assertions.assertThat;

/**
 * Image export (exportStringAsPicture) tests.
 * <p>
 * Uses the file: URL of a <b>locally generated PNG file</b> as input rather than a network resource.
 * Verifies that images are actually embedded in the exported workbook and that invalid URLs are skipped without exception.
 * <p>
 * The last group calls {@link PxlCellUtils#addPicturesToCell} directly to check the anchor coordinates each file
 * format expects, since the anchor unit differs between XLS and XLSX and only the placement can tell them apart.
 */
public class PxlPictureExportTests {

    private static Pxl pxl;

    @BeforeAll
    public static void setUpBeforeClass() {
        pxl = new Pxl();
    }

    // Creates a small PNG filled with the given color and returns its file: URL string. (Different colors yield different bytes, so they are not deduplicated.)
    private static String createPng(final Path dir, final String name, final Color color) throws Exception {
        final BufferedImage image = new BufferedImage(32, 32, BufferedImage.TYPE_INT_RGB);
        final Graphics2D g = image.createGraphics();
        g.setColor(color);
        g.fillRect(0, 0, 32, 32);
        g.dispose();

        final File file = dir.resolve(name).toFile();
        ImageIO.write(image, "png", file);
        return file.toURI().toURL().toString();
    }

    // ------------------------------------------------------------------
    // Single image + collection (multiple images) embedding
    // ------------------------------------------------------------------

    @Test
    public void exportPicture_singleAndCollection_embedded(@TempDir final Path dir) throws Exception {
        final String photoUrl = createPng(dir, "photo.png", Color.RED);
        final List<String> galleryUrls = Arrays.asList(
                createPng(dir, "g1.png", Color.GREEN),
                createPng(dir, "g2.png", Color.BLUE),
                createPng(dir, "g3.png", Color.YELLOW));

        final PictureRow row = new PictureRow();
        row.setPhoto(photoUrl);
        row.setGallery(galleryUrls);

        final Workbook workbook = pxl.exportExcel()
                .sheet(PictureRow.class, Arrays.asList(row), "Pictures")
                .override(noValidationOption())
                .toWorkbook();
        try {
            final List<? extends PictureData> pictures = workbook.getAllPictures();
            // 1 single + 3 gallery = 4 (different colors, so not deduplicated)
            assertThat(pictures).hasSize(4);
            // Embedded images are written as PNG.
            assertThat(pictures).allMatch(p -> p.getPictureType() == Workbook.PICTURE_TYPE_PNG);
        } finally {
            workbook.close();
        }
    }

    // ------------------------------------------------------------------
    // Invalid image URLs are skipped without exception (only valid ones embedded)
    // ------------------------------------------------------------------

    @Test
    public void exportPicture_invalidUrl_skipped(@TempDir final Path dir) throws Exception {
        final String validUrl = createPng(dir, "valid.png", Color.MAGENTA);
        final String missingUrl = dir.resolve("does-not-exist.png").toUri().toURL().toString();

        final PictureRow row = new PictureRow();
        row.setPhoto(missingUrl);                                   // nonexistent file -> skipped
        row.setGallery(Arrays.asList(validUrl, missingUrl));        // 1 valid + 1 invalid

        final Workbook workbook = pxl.exportExcel()
                .sheet(PictureRow.class, Arrays.asList(row), "Pictures")
                .override(noValidationOption())
                .toWorkbook();
        try {
            // Exports without exception, and only the 1 valid image is embedded.
            final List<? extends PictureData> pictures = workbook.getAllPictures();
            assertThat(pictures).hasSize(1);
        } finally {
            workbook.close();
        }
    }

    @Test
    public void exportPicture_nullValue_writesNullStringWithoutPicture() throws Exception {
        // A null value is resolved to exportNullString before any codec runs, so a picture column with nothing in it
        // never reaches the picture path - there is no location to load and nothing to skip.
        final PictureRow row = new PictureRow();
        row.setPhoto(null);
        row.setGallery(null);

        final Workbook workbook = pxl.exportExcel()
                .sheet(PictureRow.class, Arrays.asList(row), "Pictures")
                .override(noValidationOption())
                .toWorkbook();
        try {
            assertThat(workbook.getAllPictures()).isEmpty();
        } finally {
            workbook.close();
        }
    }

    // ------------------------------------------------------------------
    // Which form a value starting with '=' is written in (issue L7)
    // ------------------------------------------------------------------

    @Test
    public void exportPicture_collectionFirstElementStartingWithEquals_embedsValidElement(@TempDir final Path dir) throws Exception {
        // The Collection path never tested for a leading '=' the way the String path used to, and it still does not:
        // an element that looks like a formula is just another location that fails to load, so the valid element
        // beside it is embedded regardless. Putting it first makes the joined cell string start with '=', which is
        // exactly what used to divert the String path into text.
        final String validUrl = createPng(dir, "valid.png", Color.CYAN);

        final PictureRow row = new PictureRow();
        row.setGallery(Arrays.asList("=SUM(1,2)", validUrl));

        final Workbook workbook = pxl.exportExcel()
                .sheet(PictureRow.class, Arrays.asList(row), "Pictures")
                .override(noValidationOption())
                .toWorkbook();
        try {
            assertThat(workbook.getAllPictures()).hasSize(1);
        } finally {
            workbook.close();
        }
    }

    @Test
    public void exportPicture_valueStartingWithEquals_takesPicturePath() throws Exception {
        // The column's options decide the form, so a picture column stays a picture column even when the value looks
        // like a formula. It used to be diverted into the text path and written quote-prefixed, which no
        // exportStringAsPicture column had asked for. The location is not loadable here, so it is skipped the way any
        // bad image location is, leaving the cell empty rather than carrying the text.
        final PicturePrecedenceRow row = new PicturePrecedenceRow();
        row.setPictureOnly("=SUM(1,2)");
        row.setPictureAndFormula("=1+2");

        final Workbook workbook = pxl.exportExcel()
                .sheet(PicturePrecedenceRow.class, Arrays.asList(row), "Precedence")
                .override(noValidationOption())
                .toWorkbook();
        try {
            final Cell cell = workbook.getSheet("Precedence").getRow(1).getCell(0);
            assertThat(cell.getCellType()).isEqualTo(CellType.BLANK);
            assertThat(cell.getCellStyle().getQuotePrefixed()).isFalse();
        } finally {
            workbook.close();
        }
    }

    // ------------------------------------------------------------------
    // Anchor coordinates per file format (XLSX = EMU, XLS = fraction of the anchored cell)
    // ------------------------------------------------------------------

    @Test
    public void exportPicture_perEngine_writesFileWithEmbeddedPictures(final TestInfo testInfo, @TempDir final Path dir) throws Exception {
        // The whole picture path end to end, once per engine, leaving both files under target/test-outputs so the
        // placement can be looked at: the anchor unit differs between the formats and only opening them tells
        // whether the grid landed where it should.
        final String photoUrl = createPng(dir, "photo.png", Color.RED);
        final List<String> galleryUrls = Arrays.asList(
                createPng(dir, "g1.png", Color.GREEN),
                createPng(dir, "g2.png", Color.BLUE),
                createPng(dir, "g3.png", Color.YELLOW));

        final PictureRow row = new PictureRow();
        row.setPhoto(photoUrl);
        row.setGallery(galleryUrls);

        final File xlsFile = TestPaths.exportFile(testInfo, ".xls");
        pxl.exportExcel()
                .sheet(PictureRow.class, Arrays.asList(row), "Pictures")
                .override(PxlExportWorkbookOption.builder()
                        .exportExcelEngine(PxlExcelEngine.HSSF)
                        .exportDataValidation(false)
                        .build())
                .toFile(xlsFile);

        final File xlsxFile = TestPaths.exportFile(testInfo, ".xlsx");
        pxl.exportExcel()
                .sheet(PictureRow.class, Arrays.asList(row), "Pictures")
                .override(noValidationOption())
                .toFile(xlsxFile);

        // 1 single + 3 gallery images survive the write in both formats.
        for (final File file : Arrays.asList(xlsFile, xlsxFile)) {
            assertThat(file).exists();
            try (Workbook workbook = WorkbookFactory.create(file)) {
                assertThat(workbook.getAllPictures()).hasSize(4);
            }
        }
    }

    @Test
    public void addPicturesToCell_xls_anchorsStayInsideTheirCell(@TempDir final Path dir) throws Exception {
        final List<String> urls = Arrays.asList(
                createPng(dir, "a1.png", Color.RED),
                createPng(dir, "a2.png", Color.GREEN),
                createPng(dir, "a3.png", Color.BLUE));

        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");

            // A grid of 100px pictures runs well past one column, which XLS cannot say in a single anchor:
            // its offset is a fraction of the anchored cell, not an absolute distance like the EMU XLSX uses.
            PxlCellUtils.addPicturesToCell(sheet, urls, 100, 100, 5, 0, 0, 3);

            final List<HSSFShape> shapes = ((HSSFSheet) sheet).getDrawingPatriarch().getChildren();
            assertThat(shapes).hasSize(3);

            for (final HSSFShape shape : shapes) {
                final ClientAnchor anchor = ((Picture) shape).getClientAnchor();
                assertThat(anchor.getDx1()).isBetween(0, 1023);
                assertThat(anchor.getDx2()).isBetween(0, 1023);
                assertThat(anchor.getDy1()).isBetween(0, 255);
                assertThat(anchor.getDy2()).isBetween(0, 255);
            }

            // The second picture of the row sits a whole column further right rather than carrying an offset
            // its own column cannot hold.
            final ClientAnchor first = ((Picture) shapes.get(0)).getClientAnchor();
            final ClientAnchor second = ((Picture) shapes.get(1)).getClientAnchor();
            assertThat((int) second.getCol1()).isGreaterThan(first.getCol1());
        }
    }

    @Test
    public void addPicturesToCell_xlsZeroWidthColumn_terminates(@TempDir final Path dir) throws Exception {
        final List<String> urls = Arrays.asList(createPng(dir, "z1.png", Color.ORANGE));

        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");
            sheet.setColumnWidth(0, 0);
            sheet.setColumnWidth(1, 0);

            // A column of zero width consumes none of the offset, so the walk has to step over it instead of
            // subtracting nothing forever.
            PxlCellUtils.addPicturesToCell(sheet, urls, 100, 100, 5, 0, 0, 3);

            final List<HSSFShape> shapes = ((HSSFSheet) sheet).getDrawingPatriarch().getChildren();
            assertThat(shapes).hasSize(1);

            final ClientAnchor anchor = ((Picture) shapes.get(0)).getClientAnchor();
            assertThat((int) anchor.getCol1()).isGreaterThanOrEqualTo(2);
            assertThat(anchor.getDx1()).isBetween(0, 1023);
        }
    }

    @Test
    public void addPicturesToCell_xlsx_keepsEmuOffsets(@TempDir final Path dir) throws Exception {
        final List<String> urls = Arrays.asList(createPng(dir, "x1.png", Color.PINK));

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            final Sheet sheet = workbook.createSheet("S");

            PxlCellUtils.addPicturesToCell(sheet, urls, 100, 100, 5, 0, 0, 3);

            final List<XSSFShape> shapes = ((XSSFSheet) sheet).getDrawingPatriarch().getShapes();
            assertThat(shapes).hasSize(1);

            // XLSX measures the offset in EMU from the anchored column, so the picture stays on that column
            // and the padding is carried as an absolute distance.
            final ClientAnchor anchor = ((Picture) shapes.get(0)).getClientAnchor();
            assertThat((int) anchor.getCol1()).isZero();
            assertThat(anchor.getRow1()).isZero();
            assertThat(anchor.getDx1()).isEqualTo(Units.pixelToEMU(5));
            assertThat(anchor.getDy1()).isEqualTo(Units.pixelToEMU(5));
        }
    }

    @Test
    public void exportPicture_bothPictureAndFormula_formulaWins() throws Exception {
        // With both options set the formula wins, so the cell holds the formula rather than an image location.
        final PicturePrecedenceRow row = new PicturePrecedenceRow();
        row.setPictureOnly("=SUM(1,2)");
        row.setPictureAndFormula("=1+2");

        final Workbook workbook = pxl.exportExcel()
                .sheet(PicturePrecedenceRow.class, Arrays.asList(row), "Precedence")
                .override(noValidationOption())
                .toWorkbook();
        try {
            final Cell cell = workbook.getSheet("Precedence").getRow(1).getCell(1);
            assertThat(cell.getCellType()).isEqualTo(CellType.FORMULA);
            assertThat(cell.getCellFormula()).isEqualTo("1+2");
        } finally {
            workbook.close();
        }
    }
}
