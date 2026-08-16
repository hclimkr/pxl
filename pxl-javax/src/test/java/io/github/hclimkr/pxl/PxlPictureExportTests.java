package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.tcdata.PicturePrecedenceRow;
import io.github.hclimkr.pxl.tcdata.PictureRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import javax.imageio.ImageIO;
import java.awt.*;
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
