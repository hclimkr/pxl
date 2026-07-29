package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import io.github.hclimkr.pxl.option.PxlImportSheetOption;
import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;
import io.github.hclimkr.pxl.tcdata.BigDataRow;
import io.github.hclimkr.pxl.tcdata.TestPaths;
import org.junit.jupiter.api.*;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Large-data streaming import tests.
 * <p>
 * Generates tens of thousands of rows with SXSSF (streaming write), then reads them with the stream reader
 * (importUsingStreamReader) and verifies that the row count and values are correct. (Replaces the deprecated invoice_code large-data test.)
 * <p>
 * Since it takes some time to run, it is separated by {@code @Tag("slow")} so it can be excluded with {@code -DexcludedGroups=slow}.
 */
@Tag("slow")
public class PxlLargeDataTests {

    private static final int ROW_COUNT = 50_000;

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

    private static List<BigDataRow> generateRows(final int count) {
        final List<BigDataRow> rows = new ArrayList<>(count);
        for (int i = 0; i < count; i++) {
            final BigDataRow row = new BigDataRow();
            row.setId(i);
            row.setName("Row" + i);
            rows.add(row);
        }
        return rows;
    }

    // ------------------------------------------------------------------
    // 50,000 rows: SXSSF export -> stream reader import
    // ------------------------------------------------------------------

    @Test
    public void importLargeXlsx_streaming_readsAllRows() throws Exception {
        // Streaming write with SXSSF (minimizes write memory)
        final PxlExportWorkbookOption exportOption = PxlExportWorkbookOption.builder()
                .exportFileFormat(PxlFileFormat.SXSSF)
                .exportDataValidation(false)
                .build();

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet(BigDataRow.class, generateRows(ROW_COUNT), "Big")
                .override(exportOption)
                .toFile(excelFile);

        // Import with the stream reader (header/data row indices must be specified explicitly)
        final PxlImportSheetOption sheetOption = PxlImportSheetOption.builder()
                .importHeaderRowIndex(1)
                .importFirstDataRowIndex(2)
                .build();
        final PxlImportWorkbookOption importOption = PxlImportWorkbookOption.builder()
                .importUsingStreamReader(true)
                .importSheetOptions(Arrays.asList(sheetOption))
                .build();

        final List<BigDataRow> rows = pxl.importExcel()
                .override(importOption)
                .sheet(BigDataRow.class, Arrays.asList("Big"))
                .fromFile(excelFile);

        // Correct row count
        assertThat(rows).hasSize(ROW_COUNT);
        // Spot-check first/middle/last values
        assertThat(rows.get(0).getId()).isEqualTo(0);
        assertThat(rows.get(0).getName()).isEqualTo("Row0");
        assertThat(rows.get(ROW_COUNT / 2).getId()).isEqualTo(ROW_COUNT / 2);
        assertThat(rows.get(ROW_COUNT / 2).getName()).isEqualTo("Row" + (ROW_COUNT / 2));
        assertThat(rows.get(ROW_COUNT - 1).getId()).isEqualTo(ROW_COUNT - 1);
        assertThat(rows.get(ROW_COUNT - 1).getName()).isEqualTo("Row" + (ROW_COUNT - 1));
    }

    // ------------------------------------------------------------------
    // Whether the same large file also reads correctly non-streaming (cross-check)
    // ------------------------------------------------------------------

    @Test
    public void importLargeXlsx_nonStreaming_readsAllRows() throws Exception {
        final PxlExportWorkbookOption exportOption = PxlExportWorkbookOption.builder()
                .exportFileFormat(PxlFileFormat.SXSSF)
                .exportDataValidation(false)
                .build();

        final File excelFile = TestPaths.exportFile(testInfo);
        pxl.exportExcel()
                .sheet(BigDataRow.class, generateRows(ROW_COUNT), "Big")
                .override(exportOption)
                .toFile(excelFile);

        final List<BigDataRow> rows = pxl.importExcel()
                .sheet(BigDataRow.class, Arrays.asList("Big"))
                .fromFile(excelFile);

        assertThat(rows).hasSize(ROW_COUNT);
        assertThat(rows.get(ROW_COUNT - 1).getName()).isEqualTo("Row" + (ROW_COUNT - 1));
    }
}
