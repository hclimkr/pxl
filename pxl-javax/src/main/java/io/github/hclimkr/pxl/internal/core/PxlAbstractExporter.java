package io.github.hclimkr.pxl.internal.core;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.internal.meta.PxlExportSheetMeta;

/**
 * Format-neutral export common routine
 */
abstract class PxlAbstractExporter extends PxlAbstractBinder {

    /**
     * Resolves the sheet's row coordinates and stores them on the sheet meta.
     *
     * <p>The three {@code export*RowIndex} values are declared 1-based and are converted to the 0-based indices the
     * writers work with: the header row, the first data row (inclusive) and the data bound (exclusive). A value
     * left unspecified falls back to the row that naturally follows, and a specified one is clamped so it cannot
     * precede the level before it - the first data row never lands on or above the header, and the bound never
     * reaches past the rows actually on hand.</p>
     *
     * <p>Shared by the Excel and CSV exporters, which differ in what they do with the coordinates rather than in
     * how the coordinates are read.</p>
     *
     * @param sheetMeta    the sheet meta carrying the declared indices and receiving the resolved ones
     * @param numOfObjects the number of row objects available to write
     * @return the resolved data bound (0-based, exclusive)
     */
    protected static int resolveExportRowIndices(final PxlExportSheetMeta sheetMeta,
                                                 final int numOfObjects) {

        final int defaultHeaderRowIndex = 0;

        int actualExportHeaderRowIndex = sheetMeta.getExportHeaderRowIndex();
        if (actualExportHeaderRowIndex == PxlConstants.DEFAULT_EXPORT_HEADER_ROW_INDEX) {
            actualExportHeaderRowIndex = defaultHeaderRowIndex;
        } else {
            actualExportHeaderRowIndex -= 1;  // Specified as 1-based, so convert to 0-based.
            actualExportHeaderRowIndex = Math.max(actualExportHeaderRowIndex, defaultHeaderRowIndex);
        }

        int actualExportOriginDataRowIndex = sheetMeta.getExportFirstDataRowIndex();
        if (actualExportOriginDataRowIndex == PxlConstants.DEFAULT_EXPORT_FIRST_DATA_ROW_INDEX) {
            actualExportOriginDataRowIndex = actualExportHeaderRowIndex + 1;
        } else {
            actualExportOriginDataRowIndex -= 1;  // Specified as 1-based, so convert to 0-based.
            actualExportOriginDataRowIndex = Math.max(actualExportOriginDataRowIndex, actualExportHeaderRowIndex + 1);
        }

        int actualExportBoundDataRowIndex = sheetMeta.getExportLastDataRowIndex();
        if (actualExportBoundDataRowIndex == PxlConstants.DEFAULT_EXPORT_LAST_DATA_ROW_INDEX) {
            actualExportBoundDataRowIndex = actualExportOriginDataRowIndex + numOfObjects;
        } else {
            actualExportBoundDataRowIndex -= 1;  // Specified as 1-based, so convert to 0-based.
            actualExportBoundDataRowIndex += 1;  // Add 1 to use it as an exclusive bound.
            actualExportBoundDataRowIndex = Math.min(actualExportBoundDataRowIndex, actualExportOriginDataRowIndex + numOfObjects);
        }

        sheetMeta.setActualExportHeaderRowIndex(actualExportHeaderRowIndex);
        sheetMeta.setActualExportOriginDataRowIndex(actualExportOriginDataRowIndex);
        sheetMeta.setActualExportBoundDataRowIndex(actualExportBoundDataRowIndex);

        return actualExportBoundDataRowIndex;
    }

}
