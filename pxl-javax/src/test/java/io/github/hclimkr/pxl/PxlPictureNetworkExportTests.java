package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.tcdata.PictureRow;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;

import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Arrays;
import java.util.List;

import static io.github.hclimkr.pxl.tcdata.Fixtures.noValidationOption;
import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

/**
 * Network image export (exportStringAsPicture) tests.
 * <p>
 * Uses free image URLs from https://picsum.photos as input.
 * If the network/service is unreachable, the test is <b>skipped rather than failed</b> (assumeTrue),
 * and it is separated by {@code @Tag("network")} so it can be excluded with {@code -DexcludedGroups=network}.
 */
@Tag("network")
public class PxlPictureNetworkExportTests {

    private static Pxl pxl;

    // Uses fixed ids so the images differ (preventing deduplication).
    private static final String PHOTO_URL = "https://picsum.photos/id/237/120/120";
    private static final List<String> GALLERY_URLS = Arrays.asList(
            "https://picsum.photos/id/238/120/120",
            "https://picsum.photos/id/239/120/120");

    @BeforeAll
    public static void setUpBeforeClass() {
        pxl = new Pxl();
    }

    // Checks whether picsum.photos is reachable. (Skips the test in offline/blocked environments.)
    private static boolean networkAvailable() {
        HttpURLConnection connection = null;
        try {
            connection = (HttpURLConnection) new URL("https://picsum.photos/1").openConnection();
            connection.setConnectTimeout(4000);
            connection.setReadTimeout(4000);
            connection.setRequestMethod("GET");
            final int code = connection.getResponseCode();
            return code >= 200 && code < 400;
        } catch (Exception e) {
            return false;
        } finally {
            if (connection != null) {
                connection.disconnect();
            }
        }
    }

    @Test
    public void exportPicture_networkUrl_embedded() throws Exception {
        assumeTrue(networkAvailable(), "skipping test because the network (picsum.photos) is not reachable");

        final PictureRow row = new PictureRow();
        row.setPhoto(PHOTO_URL);
        row.setGallery(GALLERY_URLS);

        final Workbook workbook = pxl.exportExcel()
                .sheet(PictureRow.class, Arrays.asList(row), "Pictures")
                .override(noValidationOption())
                .toWorkbook();
        try {
            final List<? extends PictureData> pictures = workbook.getAllPictures();
            // 1 single + 2 gallery = 3 (different ids, so not deduplicated)
            assertThat(pictures).hasSize(3);
            // Embedded images are thumbnailed and written as PNG (even if the original is JPEG).
            assertThat(pictures).allMatch(p -> p.getPictureType() == Workbook.PICTURE_TYPE_PNG);
        } finally {
            workbook.close();
        }
    }
}
