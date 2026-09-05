package io.github.hclimkr.pxl.internal.core;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * POI event-model ({@link XSSFSheetXMLHandler.SheetContentsHandler}) callback that accumulates sheet
 * contents into an in-memory header row plus a list of data rows, padding trailing/gap cells with empty
 * strings so every row has a consistent column count. Retained only for reference and no longer used.
 *
 * @see <a href="https://blog.naver.com/PostView.nhn?blogId=tmondev&logNo=221505398958">reference article</a>
 * @deprecated superseded by the streaming/user-model importers
 */
@Deprecated
public final class PxlContentsHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

    // The data part excluding the header
    private List<List<String>> rows = new ArrayList<>();

    // The 1-row List accumulated on each cell invocation
    private List<String> row = new ArrayList<>();

    // Holds the header information
    private List<String> header = new ArrayList<>();

    // The cell number used to check for empty values
    private int currentCol = -1;

    /**
     * Resets the current-column tracker at the start of a row.
     *
     * @param rowNum the 0-based row number being started
     */
    @Override
    public void startRow(int rowNum) {

        // Initial setup value used to check for empty values
        currentCol = -1;
    }

    /**
     * Finalizes a row: row 0 becomes the header, later rows are padded to the header width and appended to the data rows.
     *
     * @param rowNum the 0-based row number being ended
     */
    @Override
    public void endRow(int rowNum) {

        if (rowNum == 0) {
            header = new ArrayList<>(row);
        } else {
            // If the header is longer than the current row, the trailing cells are empty, so pad with that many blanks
            if (row.size() < header.size()) {
                row.addAll(Collections.nCopies(header.size() - row.size(), StringUtils.EMPTY));
            }

            rows.add(new ArrayList<>(row));
        }

        row.clear();
    }

    /**
     * Appends one cell's value to the current row, filling any skipped (empty) columns with empty strings.
     *
     * @param cellReference  the A1-style reference of the cell
     * @param formattedValue the formatted cell value
     * @param comment        the cell comment, if any
     */
    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {

        int iCol = (new CellReference(cellReference)).getCol();
        int emptyCol = iCol - currentCol - 1;

        // Using the read cell's number, force-store empty values in the empty cell positions
        if (emptyCol > 0) {
            row.addAll(Collections.nCopies(emptyCol, StringUtils.EMPTY));
        }

        currentCol = iCol;
        row.add(formattedValue);
    }

    /**
     * Ignores page header/footer content.
     *
     * @param text     the header/footer text
     * @param isHeader {@code true} for a header, {@code false} for a footer
     * @param tagName  the header/footer tag name
     */
    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {
    }

    /**
     * No-op at end of sheet.
     */
    @Override
    public void endSheet() {
    }

}
