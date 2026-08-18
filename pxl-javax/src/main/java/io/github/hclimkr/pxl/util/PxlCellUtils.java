package io.github.hclimkr.pxl.util;

import com.github.pjfanning.xlsx.impl.StreamingSheet;
import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.type.PxlFileFormat;
import net.coobird.thumbnailator.Thumbnails;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.BooleanUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.sl.usermodel.PictureData.PictureType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.ImageUtils;
import org.apache.poi.util.Units;
import org.imgscalr.Scalr;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.BufferedInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

/**
 * Cell-level POI helpers: locating or creating a cell, reading it as text, writing a value of any supported type,
 * and decorating it.
 * <p>
 * The {@code setCellValue} family spans the JDK types PXL binds (numbers, {@link String}, {@code boolean},
 * {@link Date} and {@code java.time}, ...), so a caller need not pick the right POI setter per type, and
 * {@code getCellStringValue} renders a cell the way the spreadsheet displays it by honouring its number format
 * through a {@link DataFormatter} - pass the workbook's cached formatter on hot paths, as the overload without one
 * allocates a formatter per call. Also here: formula, error and blank cells; {@code getCellWithMerges}, which reads the
 * value a merged region carries from any of its cells; cell-style cloning; notes; and pictures anchored to a cell.
 * <p>
 * The lookups are null-safe: a {@code null} or streaming sheet, or an absent row/cell that is not to be created,
 * yields {@code null} rather than an exception, and the decorating methods no-op on a {@code null} cell. Since a
 * streaming sheet reports neither merged regions nor arbitrary rows, {@code getCellWithMerges} falls back to the
 * cell itself there.
 */
public final class PxlCellUtils {

    private static final Logger LOGGER = LoggerFactory.getLogger(PxlCellUtils.class);

    /**
     * The margin left around a note so that it does not sit flush against its cell's border, in pixels.
     */
    private static final int NOTE_ANCHOR_INSET_IN_PIXELS = 10;

    /**
     * An XLS (Escher) anchor states its x-offset as a fraction of the start column's width, in 1/1024 units,
     * so the largest offset that still falls inside that column is one below the divisor. XLSX measures the
     * same offset in EMU instead, as an absolute distance that may run past the column.
     * <p>
     * POI reads anchors back with the same divisor - {@code ImageUtils.WIDTH_UNITS} - but keeps it private,
     * hence this copy.
     */
    private static final int XLS_ANCHOR_DX_PER_COLUMN = 1024;

    /**
     * An XLS (Escher) anchor states its y-offset as a fraction of the start row's height, in 1/256 units.
     * POI's private {@code ImageUtils.HEIGHT_UNITS} holds the same value.
     *
     * @see #XLS_ANCHOR_DX_PER_COLUMN
     */
    private static final int XLS_ANCHOR_DY_PER_ROW = 256;

    /**
     * Prevents instantiation.
     */
    private PxlCellUtils() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Determines whether the cell is blank. A {@code null} cell, or a cell whose type is
     * {@link CellType#BLANK}, is considered blank.
     *
     * @param cell the cell to inspect, may be {@code null}
     * @return {@code true} if the cell is {@code null} or of blank type
     */
    public static boolean isBlankCell(final Cell cell) {

        return (Objects.isNull(cell) || cell.getCellType() == CellType.BLANK);
    }

    /**
     * Returns the cell at the given row/column, optionally creating the row and cell when absent.
     * Returns {@code null} if the row is unavailable (for example a {@code null} or streaming sheet, or a
     * missing row with {@code createIfNone} {@code false}) or if the cell is missing and not to be created.
     *
     * @param sheet        the sheet to read from
     * @param rowIndex     the zero-based row index
     * @param columnIndex  the zero-based column index
     * @param createIfNone whether to create the row and cell if they do not yet exist
     * @return the (possibly newly created) cell, or {@code null} if unavailable
     */
    public static Cell getCell(final Sheet sheet,
                               final int rowIndex,
                               final int columnIndex,
                               final boolean createIfNone) {

        final Row row = PxlRowUtils.getRow(sheet, rowIndex, createIfNone);
        if (Objects.nonNull(row)) {
            Cell cell = row.getCell(columnIndex);
            if (Objects.isNull(cell) && createIfNone) {
                cell = row.createCell(columnIndex);
            }

            return cell;
        }

        return null;
    }

    /**
     * Returns the cell addressed by an A1-style reference (for example {@code "B3"}), optionally
     * creating it when absent. The reference is resolved to row/column indexes and delegated to
     * {@link #getCell(Sheet, int, int, boolean)}.
     *
     * @param sheet        the sheet to read from
     * @param cellRefStr   the A1-style cell reference
     * @param createIfNone whether to create the row and cell if they do not yet exist
     * @return the (possibly newly created) cell, or {@code null} if unavailable
     * @throws PxlArgumentException if the reference does not include both a row and a column
     */
    public static Cell getCell(final Sheet sheet,
                               final String cellRefStr,
                               final boolean createIfNone)
            throws PxlArgumentException {

        final Pair<Integer, Integer> cellIndexes = PxlMiscUtils.convertCellReferenceStringToIndexes(cellRefStr);
        final int rowIndex = cellIndexes.getLeft();
        final int columnIndex = cellIndexes.getRight();

        return getCell(sheet, rowIndex, columnIndex, createIfNone);
    }

    /**
     * Copies the content of one cell onto another. The style is reused directly when both cells belong
     * to the same workbook, otherwise it is cloned into the destination workbook via
     * {@link #cloneCellStyle(Cell, Workbook)}. Any cell comment and hyperlink are re-created as new
     * objects owned by the destination sheet rather than handed over as-is, so the source keeps its own
     * and a template cell can be copied to many destinations; a comment's rich-text runs are flattened
     * to plain text in the process. The typed value is transferred according to the source cell type
     * (numeric, string, boolean, blank, formula or error). A {@code null} source or destination cell is
     * a no-op.
     *
     * @param srcCell the cell to copy from, may be {@code null}
     * @param dstCell the cell to copy to, may be {@code null}
     */
    public static void copyCell(final Cell srcCell, final Cell dstCell) {

        if (Objects.isNull(srcCell) || Objects.isNull(dstCell)) {
            return;
        }

        CellStyle cellStyle;
        if (srcCell.getSheet().getWorkbook() == dstCell.getSheet().getWorkbook()) {
            cellStyle = srcCell.getCellStyle();
        } else {
            cellStyle = PxlCellUtils.cloneCellStyle(srcCell, dstCell.getSheet().getWorkbook());
        }

        if (Objects.nonNull(cellStyle)) {
            dstCell.setCellStyle(cellStyle);
        }

        copyCellComment(srcCell, dstCell);
        copyCellHyperlink(srcCell, dstCell);

        switch (srcCell.getCellType()) {
            case NUMERIC:
                dstCell.setCellValue(srcCell.getNumericCellValue());
                break;
            case STRING:
                dstCell.setCellValue(srcCell.getRichStringCellValue());
                break;
            case BOOLEAN:
                dstCell.setCellValue(srcCell.getBooleanCellValue());
                break;
            case BLANK:
                dstCell.setBlank();
                break;
            case FORMULA:
                dstCell.setCellFormula(srcCell.getCellFormula());
                break;
            case ERROR:
                dstCell.setCellErrorValue(srcCell.getErrorCellValue());
                break;
        }
    }

    /**
     * Copies a cell to another position within the same sheet, resolving both cells by index and
     * delegating to {@link #copyCell(Cell, Cell)}. The source cell is looked up without creation, since
     * copying from a cell that does not exist has nothing to read; {@code createIfNone} governs the
     * destination, which is created when absent only if it is {@code true}. Returns {@code null} when
     * the source cell does not exist, or when the destination is absent and not to be created.
     *
     * @param sheet          the sheet to operate on
     * @param srcRowIndex    the zero-based source row index
     * @param srcColumnIndex the zero-based source column index
     * @param dstRowIndex    the zero-based destination row index
     * @param dstColumnIndex the zero-based destination column index
     * @param createIfNone   whether to create the destination row and cell if they do not yet exist
     * @return the destination cell, or {@code null} if either cell is unavailable
     */
    public static Cell copyCell(final Sheet sheet,
                                final int srcRowIndex,
                                final int srcColumnIndex,
                                final int dstRowIndex,
                                final int dstColumnIndex,
                                final boolean createIfNone) {

        final Cell srcCell = getCell(sheet, srcRowIndex, srcColumnIndex, false);
        if (Objects.isNull(srcCell)) {
            return null;
        }

        final Cell dstCell = getCell(sheet, dstRowIndex, dstColumnIndex, createIfNone);
        if (Objects.isNull(dstCell)) {
            return null;
        }

        copyCell(srcCell, dstCell);

        return dstCell;
    }

    /**
     * Copies the source cell's comment onto the destination cell as a new comment owned by the destination
     * sheet. Handing POI the source's own comment object instead would relocate it: the note would vanish
     * from the source cell, only one destination would keep it when a template cell is copied repeatedly,
     * and across workbooks it would not attach to the destination at all while moving inside the source
     * workbook. The copy carries the source's text, author and visibility, and keeps the source anchor's
     * size - at least one column and one row wide, and one cell wide when the source reports no anchor -
     * while sitting over the destination cell; rich-text runs are flattened to plain text. A source cell
     * without a comment is a no-op.
     *
     * @param srcCell the cell to read the comment from
     * @param dstCell the cell to attach the copied comment to
     */
    private static void copyCellComment(final Cell srcCell,
                                        final Cell dstCell) {

        final Comment srcComment = srcCell.getCellComment();
        if (Objects.isNull(srcComment)) {
            return;
        }

        final Sheet dstSheet = dstCell.getSheet();
        final CreationHelper creationHelper = dstSheet.getWorkbook().getCreationHelper();

        Drawing<?> drawing = dstSheet.getDrawingPatriarch();
        if (Objects.isNull(drawing)) {
            drawing = dstSheet.createDrawingPatriarch();
        }

        final ClientAnchor srcAnchor = srcComment.getClientAnchor();
        final int columnSpan = Objects.isNull(srcAnchor) ? 1 : Math.max(1, srcAnchor.getCol2() - srcAnchor.getCol1());
        final int rowSpan = Objects.isNull(srcAnchor) ? 1 : Math.max(1, srcAnchor.getRow2() - srcAnchor.getRow1());

        final ClientAnchor dstAnchor = creationHelper.createClientAnchor();
        dstAnchor.setCol1(dstCell.getColumnIndex());
        dstAnchor.setCol2(dstCell.getColumnIndex() + columnSpan);
        dstAnchor.setRow1(dstCell.getRowIndex());
        dstAnchor.setRow2(dstCell.getRowIndex() + rowSpan);

        final RichTextString srcString = srcComment.getString();
        final Comment dstComment = drawing.createCellComment(dstAnchor);
        dstComment.setAuthor(StringUtils.defaultString(srcComment.getAuthor()));
        dstComment.setString(creationHelper.createRichTextString(Objects.isNull(srcString) ? StringUtils.EMPTY : StringUtils.defaultString(srcString.getString())));
        dstComment.setVisible(srcComment.isVisible());

        dstCell.setCellComment(dstComment);
    }

    /**
     * Copies the source cell's hyperlink onto the destination cell as a new link owned by the destination
     * workbook. Handing POI the source's own hyperlink object instead would re-point it at the destination,
     * dropping it from the source cell and registering the very same object twice on the sheet, and across
     * workbooks it would leave both cells sharing one link. Type, address and label are carried over. A
     * source cell without a hyperlink is a no-op.
     *
     * @param srcCell the cell to read the hyperlink from
     * @param dstCell the cell to attach the copied hyperlink to
     */
    private static void copyCellHyperlink(final Cell srcCell,
                                          final Cell dstCell) {

        final Hyperlink srcHyperlink = srcCell.getHyperlink();
        if (Objects.isNull(srcHyperlink)) {
            return;
        }

        final CreationHelper creationHelper = dstCell.getSheet().getWorkbook().getCreationHelper();
        final Hyperlink dstHyperlink = creationHelper.createHyperlink(srcHyperlink.getType());

        if (Objects.nonNull(srcHyperlink.getAddress())) {
            dstHyperlink.setAddress(srcHyperlink.getAddress());
        }

        if (Objects.nonNull(srcHyperlink.getLabel())) {
            dstHyperlink.setLabel(srcHyperlink.getLabel());
        }

        dstCell.setHyperlink(dstHyperlink);
    }

    /**
     * Returns the cell value as a string using a default {@link DataFormatter}, by delegating to
     * {@link #getCellStringValue(Cell, DataFormatter)} with a {@code null} formatter.
     *
     * @param cell the cell to read, may be {@code null}
     * @return the formatted string value, or {@code null} if the cell is {@code null}
     */
    public static String getCellStringValue(final Cell cell) {

        return getCellStringValue(cell, null);
    }

    /**
     * Returns the cell value as a string. Numeric cells are rendered with the supplied
     * {@link DataFormatter} (or a default {@link Locale#ROOT} one when {@code null}, so decimal/grouping
     * symbols are locale-independent); streaming numeric cells are rendered the same way, since the streaming
     * reader reads styles by default and the cell carries its number format. String cells return their text,
     * boolean cells return {@code "true"}/{@code "false"}, and formula cells return the formula text. Blank
     * cells and a {@code null} cell return {@code null}.
     *
     * @param cell          the cell to read, may be {@code null}
     * @param dataFormatter the formatter for numeric cells, or {@code null} to use a default {@link Locale#ROOT} formatter
     * @return the string representation of the cell value, or {@code null} if the cell is {@code null} or blank
     */
    public static String getCellStringValue(final Cell cell,
                                            final DataFormatter dataFormatter) {

        if (Objects.isNull(cell)) {
            return null;
        }

        String stringValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                // #1 Render by applying the cell's display format. The default formatter is Locale.ROOT, so the decimal/grouping separators are fixed independently of locale.
                final DataFormatter formatter = Objects.nonNull(dataFormatter) ? dataFormatter : new DataFormatter(Locale.ROOT);
                stringValue = formatter.formatCellValue(cell);

//                // #2 With this approach the double held in the cell is turned into text as-is, ignoring the cell format
//                final double numericCellValue = cell.getNumericCellValue();
//                stringValue = NumberToTextConverter.toText(numericCellValue);

//                // #3: With this approach there is an issue where a cell value of 2012000046 is converted to the string "2.012000046E9".
//                final double numericCellValue = cell.getNumericCellValue();
//                stringValue = String.valueOf(numericCellValue);

                break;

            case STRING:
                stringValue = cell.getStringCellValue();
                break;

            case BOOLEAN:
                final boolean booleanCellValue = cell.getBooleanCellValue();
                stringValue = BooleanUtils.toStringTrueFalse(booleanCellValue);
                break;

            case FORMULA:
                stringValue = cell.getCellFormula();
                break;

            case BLANK:
            default:
                // empty
                break;
        }

        return stringValue;
    }

    /**
     * Sets a cell's value from an arbitrary object located by row/column. A {@code null} value blanks
     * the cell; a {@link Number} is written as a numeric double; any other value is written as its
     * {@link String#valueOf(Object)} text.
     *
     * @param sheet        the sheet to operate on
     * @param rowIndex     the zero-based row index
     * @param columnIndex  the zero-based column index
     * @param value        the value to write, may be {@code null}
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final int rowIndex,
                                    final int columnIndex,
                                    final Object value,
                                    final boolean createIfNone) {

        return Optional.ofNullable(getCell(sheet, rowIndex, columnIndex, createIfNone))
                .map(cell -> {
                    if (Objects.isNull(value)) {
                        cell.setBlank();
                    } else {
                        if (value instanceof Number) {
                            cell.setCellValue(((Number) value).doubleValue());
                        } else {
                            cell.setCellValue(String.valueOf(value));
                        }
                    }
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a cell's value from an arbitrary object located by an A1-style reference. A {@code null}
     * value blanks the cell; a {@link Number} is written as a numeric double; any other value is written
     * as its {@link String#valueOf(Object)} text.
     *
     * @param sheet        the sheet to operate on
     * @param cellRefStr   the A1-style cell reference
     * @param value        the value to write, may be {@code null}
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     * @throws PxlArgumentException if the reference does not include both a row and a column
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final String cellRefStr,
                                    final Object value,
                                    final boolean createIfNone)
            throws PxlArgumentException {

        return Optional.ofNullable(getCell(sheet, cellRefStr, createIfNone))
                .map(cell -> {
                    if (Objects.isNull(value)) {
                        cell.setBlank();
                    } else {
                        if (value instanceof Number) {
                            cell.setCellValue(((Number) value).doubleValue());
                        } else {
                            cell.setCellValue(String.valueOf(value));
                        }
                    }
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a numeric cell value located by row/column.
     *
     * @param sheet        the sheet to operate on
     * @param rowIndex     the zero-based row index
     * @param columnIndex  the zero-based column index
     * @param value        the numeric value to write
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final int rowIndex,
                                    final int columnIndex,
                                    final double value,
                                    final boolean createIfNone) {

        return Optional.ofNullable(getCell(sheet, rowIndex, columnIndex, createIfNone))
                .map(cell -> {
                    cell.setCellValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a numeric cell value located by an A1-style reference.
     *
     * @param sheet        the sheet to operate on
     * @param cellRefStr   the A1-style cell reference
     * @param value        the numeric value to write
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     * @throws PxlArgumentException if the reference does not include both a row and a column
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final String cellRefStr,
                                    final double value,
                                    final boolean createIfNone)
            throws PxlArgumentException {

        return Optional.ofNullable(getCell(sheet, cellRefStr, createIfNone))
                .map(cell -> {
                    cell.setCellValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a boolean cell value located by row/column; a {@code null} value blanks the cell.
     *
     * @param sheet        the sheet to operate on
     * @param rowIndex     the zero-based row index
     * @param columnIndex  the zero-based column index
     * @param value        the boolean value to write, or {@code null} to blank the cell
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final int rowIndex,
                                    final int columnIndex,
                                    final Boolean value,
                                    final boolean createIfNone) {

        return Optional.ofNullable(getCell(sheet, rowIndex, columnIndex, createIfNone))
                .map(cell -> {
                    if (Objects.isNull(value)) {
                        cell.setBlank();
                    } else {
                        cell.setCellValue(value);
                    }
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a boolean cell value located by an A1-style reference; a {@code null} value blanks the cell.
     *
     * @param sheet        the sheet to operate on
     * @param cellRefStr   the A1-style cell reference
     * @param value        the boolean value to write, or {@code null} to blank the cell
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     * @throws PxlArgumentException if the reference does not include both a row and a column
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final String cellRefStr,
                                    final Boolean value,
                                    final boolean createIfNone)
            throws PxlArgumentException {

        return Optional.ofNullable(getCell(sheet, cellRefStr, createIfNone))
                .map(cell -> {
                    if (Objects.isNull(value)) {
                        cell.setBlank();
                    } else {
                        cell.setCellValue(value);
                    }
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a primitive boolean cell value located by row/column.
     *
     * @param sheet        the sheet to operate on
     * @param rowIndex     the zero-based row index
     * @param columnIndex  the zero-based column index
     * @param value        the boolean value to write
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final int rowIndex,
                                    final int columnIndex,
                                    final boolean value,
                                    final boolean createIfNone) {

        return Optional.ofNullable(getCell(sheet, rowIndex, columnIndex, createIfNone))
                .map(cell -> {
                    cell.setCellValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a primitive boolean cell value located by an A1-style reference.
     *
     * @param sheet        the sheet to operate on
     * @param cellRefStr   the A1-style cell reference
     * @param value        the boolean value to write
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     * @throws PxlArgumentException if the reference does not include both a row and a column
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final String cellRefStr,
                                    final boolean value,
                                    final boolean createIfNone)
            throws PxlArgumentException {

        return Optional.ofNullable(getCell(sheet, cellRefStr, createIfNone))
                .map(cell -> {
                    cell.setCellValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a {@link Date} cell value located by row/column.
     *
     * @param sheet        the sheet to operate on
     * @param rowIndex     the zero-based row index
     * @param columnIndex  the zero-based column index
     * @param value        the date value to write
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final int rowIndex,
                                    final int columnIndex,
                                    final Date value,
                                    final boolean createIfNone) {

        return Optional.ofNullable(getCell(sheet, rowIndex, columnIndex, createIfNone))
                .map(cell -> {
                    cell.setCellValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a {@link Date} cell value located by an A1-style reference.
     *
     * @param sheet        the sheet to operate on
     * @param cellRefStr   the A1-style cell reference
     * @param value        the date value to write
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     * @throws PxlArgumentException if the reference does not include both a row and a column
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final String cellRefStr,
                                    final Date value,
                                    final boolean createIfNone)
            throws PxlArgumentException {

        return Optional.ofNullable(getCell(sheet, cellRefStr, createIfNone))
                .map(cell -> {
                    cell.setCellValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a {@link LocalDateTime} cell value located by row/column.
     *
     * @param sheet        the sheet to operate on
     * @param rowIndex     the zero-based row index
     * @param columnIndex  the zero-based column index
     * @param value        the date-time value to write
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final int rowIndex,
                                    final int columnIndex,
                                    final LocalDateTime value,
                                    final boolean createIfNone) {

        return Optional.ofNullable(getCell(sheet, rowIndex, columnIndex, createIfNone))
                .map(cell -> {
                    cell.setCellValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a {@link LocalDateTime} cell value located by an A1-style reference.
     *
     * @param sheet        the sheet to operate on
     * @param cellRefStr   the A1-style cell reference
     * @param value        the date-time value to write
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     * @throws PxlArgumentException if the reference does not include both a row and a column
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final String cellRefStr,
                                    final LocalDateTime value,
                                    final boolean createIfNone)
            throws PxlArgumentException {

        return Optional.ofNullable(getCell(sheet, cellRefStr, createIfNone))
                .map(cell -> {
                    cell.setCellValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a {@link LocalDate} cell value located by row/column.
     *
     * @param sheet        the sheet to operate on
     * @param rowIndex     the zero-based row index
     * @param columnIndex  the zero-based column index
     * @param value        the date value to write
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final int rowIndex,
                                    final int columnIndex,
                                    final LocalDate value,
                                    final boolean createIfNone) {

        return Optional.ofNullable(getCell(sheet, rowIndex, columnIndex, createIfNone))
                .map(cell -> {
                    cell.setCellValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a {@link LocalDate} cell value located by an A1-style reference.
     *
     * @param sheet        the sheet to operate on
     * @param cellRefStr   the A1-style cell reference
     * @param value        the date value to write
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     * @throws PxlArgumentException if the reference does not include both a row and a column
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final String cellRefStr,
                                    final LocalDate value,
                                    final boolean createIfNone)
            throws PxlArgumentException {

        return Optional.ofNullable(getCell(sheet, cellRefStr, createIfNone))
                .map(cell -> {
                    cell.setCellValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a {@link Calendar} cell value located by row/column.
     *
     * @param sheet        the sheet to operate on
     * @param rowIndex     the zero-based row index
     * @param columnIndex  the zero-based column index
     * @param value        the calendar value to write
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final int rowIndex,
                                    final int columnIndex,
                                    final Calendar value,
                                    final boolean createIfNone) {

        return Optional.ofNullable(getCell(sheet, rowIndex, columnIndex, createIfNone))
                .map(cell -> {
                    cell.setCellValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a {@link Calendar} cell value located by an A1-style reference.
     *
     * @param sheet        the sheet to operate on
     * @param cellRefStr   the A1-style cell reference
     * @param value        the calendar value to write
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     * @throws PxlArgumentException if the reference does not include both a row and a column
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final String cellRefStr,
                                    final Calendar value,
                                    final boolean createIfNone)
            throws PxlArgumentException {

        return Optional.ofNullable(getCell(sheet, cellRefStr, createIfNone))
                .map(cell -> {
                    cell.setCellValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a {@link RichTextString} cell value located by row/column.
     *
     * @param sheet        the sheet to operate on
     * @param rowIndex     the zero-based row index
     * @param columnIndex  the zero-based column index
     * @param value        the rich-text value to write
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final int rowIndex,
                                    final int columnIndex,
                                    final RichTextString value,
                                    final boolean createIfNone) {

        return Optional.ofNullable(getCell(sheet, rowIndex, columnIndex, createIfNone))
                .map(cell -> {
                    cell.setCellValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a {@link RichTextString} cell value located by an A1-style reference.
     *
     * @param sheet        the sheet to operate on
     * @param cellRefStr   the A1-style cell reference
     * @param value        the rich-text value to write
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     * @throws PxlArgumentException if the reference does not include both a row and a column
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final String cellRefStr,
                                    final RichTextString value,
                                    final boolean createIfNone)
            throws PxlArgumentException {

        return Optional.ofNullable(getCell(sheet, cellRefStr, createIfNone))
                .map(cell -> {
                    cell.setCellValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a {@link String} cell value located by row/column.
     *
     * @param sheet        the sheet to operate on
     * @param rowIndex     the zero-based row index
     * @param columnIndex  the zero-based column index
     * @param value        the string value to write
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final int rowIndex,
                                    final int columnIndex,
                                    final String value,
                                    final boolean createIfNone) {

        return Optional.ofNullable(getCell(sheet, rowIndex, columnIndex, createIfNone))
                .map(cell -> {
                    cell.setCellValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a {@link String} cell value located by an A1-style reference.
     *
     * @param sheet        the sheet to operate on
     * @param cellRefStr   the A1-style cell reference
     * @param value        the string value to write
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     * @throws PxlArgumentException if the reference does not include both a row and a column
     */
    public static Cell setCellValue(final Sheet sheet,
                                    final String cellRefStr,
                                    final String value,
                                    final boolean createIfNone)
            throws PxlArgumentException {

        return Optional.ofNullable(getCell(sheet, cellRefStr, createIfNone))
                .map(cell -> {
                    cell.setCellValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a cell formula located by row/column.
     *
     * @param sheet        the sheet to operate on
     * @param rowIndex     the zero-based row index
     * @param columnIndex  the zero-based column index
     * @param formula      the formula text (without a leading {@code =})
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     */
    public static Cell setCellFormula(final Sheet sheet,
                                      final int rowIndex,
                                      final int columnIndex,
                                      final String formula,
                                      final boolean createIfNone) {

        return Optional.ofNullable(getCell(sheet, rowIndex, columnIndex, createIfNone))
                .map(cell -> {
                    cell.setCellFormula(formula);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a cell formula located by an A1-style reference.
     *
     * @param sheet        the sheet to operate on
     * @param cellRefStr   the A1-style cell reference
     * @param formula      the formula text (without a leading {@code =})
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     * @throws PxlArgumentException if the reference does not include both a row and a column
     */
    public static Cell setCellFormula(final Sheet sheet,
                                      final String cellRefStr,
                                      final String formula,
                                      final boolean createIfNone)
            throws PxlArgumentException {

        return Optional.ofNullable(getCell(sheet, cellRefStr, createIfNone))
                .map(cell -> {
                    cell.setCellFormula(formula);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a cell error value located by row/column.
     *
     * @param sheet        the sheet to operate on
     * @param rowIndex     the zero-based row index
     * @param columnIndex  the zero-based column index
     * @param value        the POI error code (see {@link FormulaError})
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     */
    public static Cell setCellErrorValue(final Sheet sheet,
                                         final int rowIndex,
                                         final int columnIndex,
                                         final byte value,
                                         final boolean createIfNone) {

        return Optional.ofNullable(getCell(sheet, rowIndex, columnIndex, createIfNone))
                .map(cell -> {
                    cell.setCellErrorValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Sets a cell error value located by an A1-style reference.
     *
     * @param sheet        the sheet to operate on
     * @param cellRefStr   the A1-style cell reference
     * @param value        the POI error code (see {@link FormulaError})
     * @param createIfNone whether to create the row and cell if absent
     * @return the written cell, or {@code null} if the cell is unavailable
     * @throws PxlArgumentException if the reference does not include both a row and a column
     */
    public static Cell setCellErrorValue(final Sheet sheet,
                                         final String cellRefStr,
                                         final byte value,
                                         final boolean createIfNone)
            throws PxlArgumentException {

        return Optional.ofNullable(getCell(sheet, cellRefStr, createIfNone))
                .map(cell -> {
                    cell.setCellErrorValue(value);
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Blanks the cell located by row/column.
     *
     * @param sheet        the sheet to operate on
     * @param rowIndex     the zero-based row index
     * @param columnIndex  the zero-based column index
     * @param createIfNone whether to create the row and cell if absent
     * @return the blanked cell, or {@code null} if the cell is unavailable
     */
    public static Cell setCellBlank(final Sheet sheet,
                                    final int rowIndex,
                                    final int columnIndex,
                                    final boolean createIfNone) {

        return Optional.ofNullable(getCell(sheet, rowIndex, columnIndex, createIfNone))
                .map(cell -> {
                    cell.setBlank();
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Blanks the cell located by an A1-style reference.
     *
     * @param sheet        the sheet to operate on
     * @param cellRefStr   the A1-style cell reference
     * @param createIfNone whether to create the row and cell if absent
     * @return the blanked cell, or {@code null} if the cell is unavailable
     * @throws PxlArgumentException if the reference does not include both a row and a column
     */
    public static Cell setCellBlank(final Sheet sheet,
                                    final String cellRefStr,
                                    final boolean createIfNone)
            throws PxlArgumentException {

        return Optional.ofNullable(getCell(sheet, cellRefStr, createIfNone))
                .map(cell -> {
                    cell.setBlank();
                    return cell;
                })
                .orElse(null);
    }

    /**
     * Returns the cell at the given position, resolving merged regions to their anchor cell. If the
     * cell itself is non-blank it is returned as-is. Otherwise, when the position falls inside a merged
     * region, the region's top-left (first row/first column) cell is returned, since that is where a
     * merged region's value lives. For a streaming sheet ({@link StreamingSheet}), which cannot report
     * merged regions, the plain cell is returned without merge resolution, and a {@code null} sheet
     * yields {@code null}.
     *
     * @param sheet       the sheet to read from
     * @param rowIndex    the zero-based row index
     * @param columnIndex the zero-based column index
     * @return the value-bearing cell, or the plain (possibly {@code null}) cell if no merge applies
     */
    public static Cell getCellWithMerges(final Sheet sheet,
                                         final int rowIndex,
                                         final int columnIndex) {

        if (Objects.isNull(sheet)) {
            return null;
        }

        final Cell cell = getCell(sheet, rowIndex, columnIndex, false);

        if (sheet instanceof StreamingSheet) {
            return cell;
        }
        if (!PxlCellUtils.isBlankCell(cell)) {
            return cell;
        }

        return sheet.getMergedRegions().stream()
                .filter(cellRangeAddress -> cellRangeAddress.isInRange(rowIndex, columnIndex))
                .findFirst()
                .map(cellRangeAddress -> getCell(sheet, cellRangeAddress.getFirstRow(), cellRangeAddress.getFirstColumn(), false))
                .orElse(cell);
    }

    /**
     * Clones the cell's style into a new style within the cell's own workbook, by delegating to
     * {@link #cloneCellStyle(Cell, Workbook)}.
     *
     * @param cell the cell whose style is cloned
     * @return the newly created style, or {@code null} if the workbook's style limit is exceeded
     */
    public static CellStyle cloneCellStyle(final Cell cell) {

        return cloneCellStyle(cell, cell.getSheet().getWorkbook());
    }

    /**
     * Creates a new cell style in the target workbook copied from the given cell's style. If the target
     * workbook's maximum number of cell styles has been reached (POI raises {@link IllegalStateException}),
     * the exception is swallowed and {@code null} is returned.
     *
     * @param cell           the cell whose style is used as the source
     * @param targetWorkbook the workbook in which to create the cloned style
     * @return the newly created style, or {@code null} if the workbook's style limit is exceeded
     */
    public static CellStyle cloneCellStyle(final Cell cell,
                                           final Workbook targetWorkbook) {

        CellStyle cellStyle = null;

        try {
            cellStyle = targetWorkbook.createCellStyle();
            cellStyle.cloneStyleFrom(cell.getCellStyle());
        } catch (IllegalStateException ignored) {
        }

        return cellStyle;
    }

    /**
     * Attaches a comment (note) to the cell, anchored over the cell's own area and authored as
     * {@link PxlConstants#PXL_CREATOR}. The note is inset from the cell's border by a fixed pixel margin,
     * converted to whichever anchor unit the workbook's format takes - EMU for XLSX, a fraction of the
     * cell for XLS. A {@code null} cell or a blank note is a no-op.
     *
     * @param cell the cell to annotate
     * @param note the comment text; ignored if blank
     */
    public static void addNoteToCell(final Cell cell,
                                     final String note) {

        if (Objects.isNull(cell) || StringUtils.isBlank(note)) {
            return;
        }

        final int colIndex = cell.getColumnIndex();
        final int rowIndex = cell.getRowIndex();
        final Sheet sheet = cell.getSheet();
        final CreationHelper creationHelper = sheet.getWorkbook().getCreationHelper();
        final Drawing<?> drawing = sheet.createDrawingPatriarch();

        final RichTextString richTextString = creationHelper.createRichTextString(note);
        //TODO: need to adjust the font size
        //richTextString.applyFont(font);

        final Pair<Integer, Integer> startColumn = resolveAnchorColumn(sheet, colIndex, NOTE_ANCHOR_INSET_IN_PIXELS);
        final Pair<Integer, Integer> startRow = resolveAnchorRow(sheet, rowIndex, NOTE_ANCHOR_INSET_IN_PIXELS);
        final Pair<Integer, Integer> endColumn = resolveAnchorColumn(sheet, colIndex + 1, NOTE_ANCHOR_INSET_IN_PIXELS);
        final Pair<Integer, Integer> endRow = resolveAnchorRow(sheet, rowIndex + 1, NOTE_ANCHOR_INSET_IN_PIXELS);

        final ClientAnchor anchor = drawing.createAnchor(
                startColumn.getRight(),
                startRow.getRight(),
                endColumn.getRight(),
                endRow.getRight(),
                startColumn.getLeft(),
                startRow.getLeft(),
                endColumn.getLeft(),
                endRow.getLeft());

        final Comment comment = drawing.createCellComment(anchor);
        comment.setAuthor(PxlConstants.PXL_CREATOR);
        comment.setString(richTextString);
//        comment.setVisible(true);

        cell.setCellComment(comment);
    }

    /**
     * Adds one or more pictures over a cell, resolving the cell's position and delegating to
     * {@link #addPicturesToCell(Sheet, List, int, int, int, int, int, int)}. A {@code null} cell is a no-op.
     *
     * @param cell               the target cell whose row/column locates the pictures
     * @param imageFileUrls      the image source URLs; each is loaded and embedded
     * @param pictureWidthPx     the width of each picture in pixels
     * @param pictureHeightPx    the height of each picture in pixels
     * @param picturePaddingPx   the padding between and around pictures in pixels
     * @param horizontalImageNum the number of pictures per row before wrapping to the next line
     */
    public static void addPicturesToCell(final Cell cell,
                                         final List<String> imageFileUrls,
                                         final int pictureWidthPx,
                                         final int pictureHeightPx,
                                         final int picturePaddingPx,
                                         final int horizontalImageNum) {

        if (Objects.isNull(cell)) {
            return;
        }

        final Sheet sheet = cell.getSheet();
        final int colIndex = cell.getColumnIndex();
        final int rowIndex = cell.getRowIndex();

        addPicturesToCell(sheet, imageFileUrls, pictureWidthPx, pictureHeightPx, picturePaddingPx, colIndex, rowIndex, horizontalImageNum);
    }

    /**
     * Adds one or more pictures onto a sheet at the given column/row, laid out in a grid of
     * {@code horizontalImageNum} columns. The target row is created if absent and its height is raised
     * to fit the resulting grid. Each image is fetched from its URL and anchored with
     * {@link ClientAnchor.AnchorType#MOVE_AND_RESIZE}; an image that fails to load is skipped with a
     * warning logged via SLF4J rather than aborting the rest. A {@code null} sheet is a no-op.
     * <p>
     * The grid positions are given in pixels and converted to the anchor unit of the workbook's format:
     * XLSX takes an absolute EMU distance, while XLS takes a fraction of the anchored cell and therefore
     * has each picture anchored on the column and row its offset actually falls in.
     *
     * @param sheet              the sheet to add pictures to
     * @param imageFileUrls      the image source URLs; each is loaded and embedded
     * @param pictureWidthPx     the width of each picture in pixels
     * @param pictureHeightPx    the height of each picture in pixels
     * @param picturePaddingPx   the padding between and around pictures in pixels
     * @param colIndex           the zero-based column index the pictures are anchored on
     * @param rowIndex           the zero-based row index the pictures are anchored on
     * @param horizontalImageNum the number of pictures per row before wrapping to the next line
     */
    public static void addPicturesToCell(final Sheet sheet,
                                         final List<String> imageFileUrls,
                                         final int pictureWidthPx,
                                         final int pictureHeightPx,
                                         final int picturePaddingPx,
                                         final int colIndex,
                                         final int rowIndex,
                                         final int horizontalImageNum) {

        if (Objects.isNull(sheet)) {
            return;
        }

        Row row = sheet.getRow(rowIndex);
        if (Objects.isNull(row)) {
            row = sheet.createRow(rowIndex);
        }

        final int imageNum = PxlCollectionUtils.size(imageFileUrls);
        final int verticalImageNum = (imageNum + horizontalImageNum - 1) / horizontalImageNum;

        if (imageNum > 0) {
            final int rowHeightPx = (pictureHeightPx + picturePaddingPx) * verticalImageNum + picturePaddingPx;
            final float rowHeightInPoints = (float) Units.pixelToPoints(rowHeightPx);
            if (row.getHeightInPoints() < rowHeightInPoints) {
                row.setHeightInPoints(rowHeightInPoints);
            }
        }

        for (int imageIndex = 0; imageIndex < imageNum; imageIndex++) {

            final int horizontalIndex = imageIndex % horizontalImageNum;
            final int verticalIndex = imageIndex / horizontalImageNum;
            final int dx1 = (pictureWidthPx + picturePaddingPx) * horizontalIndex + picturePaddingPx;
            final int dy1 = (pictureHeightPx + picturePaddingPx) * verticalIndex + picturePaddingPx;
            final int dx2 = (pictureWidthPx + picturePaddingPx) * (horizontalIndex + 1);
            final int dy2 = (pictureHeightPx + picturePaddingPx) * (verticalIndex + 1);

            final String imageFileUrl = PxlCollectionUtils.get(imageFileUrls, imageIndex);
            try {
                addPictureToCell(sheet,
                        imageFileUrl,
                        ClientAnchor.AnchorType.MOVE_AND_RESIZE,
                        colIndex, rowIndex, colIndex, rowIndex,
                        dx1, dy1, dx2, dy2);
            } catch (IOException e) {
                LOGGER.warn(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.LOG_IMAGE_ADD_FAILED, String.valueOf(imageFileUrl), e.getMessage()));
            }
        }
    }

    /**
     * Loads a single image from its URL and embeds it into the sheet at the given anchor bounds. The picture bytes and
     * type are resolved via {@link #getPictureBytes(String)} and added to the workbook. A {@code null} sheet is a no-op.
     * <p>
     * The offsets are pixel distances from the edge of the cell they are counted from, and may run past it - a grid
     * of pictures anchored on one cell is laid out this way. They are converted into the unit the workbook's format
     * expects, which for XLS can move the anchor onto a later column or row; see
     * {@link #resolveAnchorColumn(Sheet, int, int)}.
     *
     * @param sheet        the sheet to add the picture to
     * @param imageFileUrl the image source URL
     * @param anchorType   the anchor behavior (for example {@link ClientAnchor.AnchorType#MOVE_AND_RESIZE})
     * @param col1         the zero-based column the start offsets are counted from
     * @param row1         the zero-based row the start offsets are counted from
     * @param col2         the zero-based column the end offsets are counted from
     * @param row2         the zero-based row the end offsets are counted from
     * @param dx1          the start x-offset from that column's left edge, in pixels
     * @param dy1          the start y-offset from that row's top edge, in pixels
     * @param dx2          the end x-offset from that column's left edge, in pixels
     * @param dy2          the end y-offset from that row's top edge, in pixels
     * @throws IOException if the image cannot be loaded
     */
    private static void addPictureToCell(final Sheet sheet,
                                         final String imageFileUrl,
                                         final ClientAnchor.AnchorType anchorType,
                                         final int col1,
                                         final int row1,
                                         final int col2,
                                         final int row2,
                                         final int dx1,
                                         final int dy1,
                                         final int dx2,
                                         final int dy2)
            throws IOException {

        if (Objects.isNull(sheet)) {
            return;
        }

        final Workbook workbook = sheet.getWorkbook();
        final CreationHelper helper = workbook.getCreationHelper();
        final ClientAnchor anchor = helper.createClientAnchor();
        anchor.setAnchorType(anchorType);

        // Place the anchor before the picture is created: an XLS anchor cannot hold an offset wider than its
        // start cell, so the resolvers may have to move the anchor onto a later column or row to express one.
        final Pair<Integer, Integer> startColumn = resolveAnchorColumn(sheet, col1, dx1);
        final Pair<Integer, Integer> startRow = resolveAnchorRow(sheet, row1, dy1);
        final Pair<Integer, Integer> endColumn = resolveAnchorColumn(sheet, col2, dx2);
        final Pair<Integer, Integer> endRow = resolveAnchorRow(sheet, row2, dy2);

        anchor.setCol1(startColumn.getLeft());
        anchor.setDx1(startColumn.getRight());
        anchor.setRow1(startRow.getLeft());
        anchor.setDy1(startRow.getRight());
        anchor.setCol2(endColumn.getLeft());
        anchor.setDx2(endColumn.getRight());
        anchor.setRow2(endRow.getLeft());
        anchor.setDy2(endRow.getRight());

        final Pair<byte[], Integer> picturePair = getPictureBytes(imageFileUrl);

        final int pictureId = workbook.addPicture(picturePair.getLeft(), picturePair.getRight());

        final Drawing drawing = sheet.createDrawingPatriarch();

        drawing.createPicture(anchor, pictureId);
    }

    /**
     * Resolves a horizontal pixel offset measured from the start column's left edge into the (column,
     * x-offset) pair the workbook's format expects.
     * <p>
     * XLSX (and its streaming variant) states the offset in EMU as an absolute distance, so the column is
     * returned unchanged and the offset may well run past its right edge - that is how a picture laid out in
     * a grid reaches beyond the cell it is anchored on. XLS instead states it as a fraction of the start
     * column's own width ({@link #XLS_ANCHOR_DX_PER_COLUMN}ths of it), which cannot express anything outside
     * that column, so the offset is walked across the columns it spans and only the remainder within the
     * column it lands in is converted. A column of zero width consumes nothing of the offset and is stepped
     * over rather than walked forever, and the walk stops at the last column the format allows.
     *
     * @param sheet            the sheet the anchor belongs to
     * @param startColumnIndex the zero-based column the offset is measured from
     * @param offsetInPixels   the offset from that column's left edge, in pixels
     * @return a pair of (zero-based column index, x-offset in the format's own unit)
     */
    private static Pair<Integer, Integer> resolveAnchorColumn(final Sheet sheet,
                                                              final int startColumnIndex,
                                                              final int offsetInPixels) {

        if (PxlFileFormat.fromPoiWorkbook(sheet.getWorkbook()) != PxlFileFormat.XLS) {
            return Pair.of(startColumnIndex, Units.pixelToEMU(offsetInPixels));
        }

        final int lastColumnIndex = PxlFileFormat.XLS.getMaxExportColumns() - 1;

        int columnIndex = Math.min(Math.max(0, startColumnIndex), lastColumnIndex);
        double remainingInPixels = Math.max(0, offsetInPixels);

        while (columnIndex < lastColumnIndex) {
            final double columnWidthInPixels = sheet.getColumnWidthInPixels(columnIndex);
            if (columnWidthInPixels > 0.D && remainingInPixels < columnWidthInPixels) {
                break;
            }

            remainingInPixels -= Math.max(0.D, columnWidthInPixels);
            columnIndex++;
        }

        return Pair.of(columnIndex, toAnchorFraction(remainingInPixels, sheet.getColumnWidthInPixels(columnIndex), XLS_ANCHOR_DX_PER_COLUMN));
    }

    /**
     * Resolves a vertical pixel offset measured from the start row's top edge into the (row, y-offset) pair
     * the workbook's format expects, the way {@link #resolveAnchorColumn(Sheet, int, int)} does for columns.
     * XLS states the offset in {@link #XLS_ANCHOR_DY_PER_ROW}ths of the start row's height. Row heights come
     * from POI's own {@code ImageUtils.getRowHeightInPixels}, which falls back to the sheet's default height
     * for a row that does not exist yet - the same measurement POI uses when it reads an anchor back.
     *
     * @param sheet          the sheet the anchor belongs to
     * @param startRowIndex  the zero-based row the offset is measured from
     * @param offsetInPixels the offset from that row's top edge, in pixels
     * @return a pair of (zero-based row index, y-offset in the format's own unit)
     */
    private static Pair<Integer, Integer> resolveAnchorRow(final Sheet sheet,
                                                           final int startRowIndex,
                                                           final int offsetInPixels) {

        if (PxlFileFormat.fromPoiWorkbook(sheet.getWorkbook()) != PxlFileFormat.XLS) {
            return Pair.of(startRowIndex, Units.pixelToEMU(offsetInPixels));
        }

        final int lastRowIndex = PxlFileFormat.XLS.getMaxExportRows() - 1;

        int rowIndex = Math.min(Math.max(0, startRowIndex), lastRowIndex);
        double remainingInPixels = Math.max(0, offsetInPixels);

        while (rowIndex < lastRowIndex) {
            final double rowHeightInPixels = ImageUtils.getRowHeightInPixels(sheet, rowIndex);
            if (rowHeightInPixels > 0.D && remainingInPixels < rowHeightInPixels) {
                break;
            }

            remainingInPixels -= Math.max(0.D, rowHeightInPixels);
            rowIndex++;
        }

        return Pair.of(rowIndex, toAnchorFraction(remainingInPixels, ImageUtils.getRowHeightInPixels(sheet, rowIndex), XLS_ANCHOR_DY_PER_ROW));
    }

    /**
     * Converts an offset that lies within one cell into the fraction of that cell an XLS anchor stores,
     * clamped to the range POI accepts (zero up to one below {@code fractionsPerCell}). A cell with no
     * measurable extent leaves no room for an offset, so it yields zero.
     *
     * @param offsetInPixels   the offset within the cell, in pixels
     * @param cellSizeInPixels the cell's width or height, in pixels
     * @param fractionsPerCell the number of fractions the cell is divided into
     * @return the offset as a fraction of the cell
     */
    private static int toAnchorFraction(final double offsetInPixels,
                                        final double cellSizeInPixels,
                                        final int fractionsPerCell) {

        if (cellSizeInPixels <= 0.D || offsetInPixels <= 0.D) {
            return 0;
        }

        return (int) Math.min(fractionsPerCell - 1L, Math.round(fractionsPerCell * offsetInPixels / cellSizeInPixels));
    }

    /**
     * Reads the image at the URL and returns its bytes paired with the POI picture type. Depending on
     * {@link PxlConstants#EXPORT_PICTURE_SCALER}, the image is either passed through unchanged (with its type probed via
     * {@link #probePictureType(String, InputStream)}) or downscaled to a PNG using Thumbnailator or imgscalr.
     *
     * @param imageFileUrl the image source URL
     * @return a pair of (image bytes, POI OOXML picture-type id)
     * @throws IOException if the image cannot be read or is not a readable image
     */
    private static Pair<byte[], Integer> getPictureBytes(final String imageFileUrl)
            throws IOException {

        final URL url = new URL(imageFileUrl);
        int pictureType;
        byte[] pictureBytes;

        switch (PxlConstants.EXPORT_PICTURE_SCALER) {
            default:
            case NO_SCALER: {
                InputStream imageInputStream = null;
                try {
                    imageInputStream = new BufferedInputStream(url.openStream());

                    pictureType = probePictureType(imageFileUrl, imageInputStream);
                    pictureBytes = IOUtils.toByteArray(imageInputStream);
                } finally {
                    IOUtils.closeQuietly(imageInputStream);
                }

                break;
            }

            case THUMBNAILATOR: {
                pictureType = PictureType.PNG.ooxmlId;  // Workbook.PICTURE_TYPE_PNG

                final BufferedImage thumbImage = Thumbnails.of(url)
                        .size(PxlConstants.EXPORT_PICTURE_SCALE_WIDTH_IN_PIXELS, PxlConstants.EXPORT_PICTURE_SCALE_HEIGHT_IN_PIXELS)
                        .asBufferedImage();

                ByteArrayOutputStream thumbByteArrayOutputStream = null;
                try {
                    thumbByteArrayOutputStream = new ByteArrayOutputStream();

                    ImageIO.write(thumbImage, "png", thumbByteArrayOutputStream);
                    pictureBytes = thumbByteArrayOutputStream.toByteArray();
                } finally {
                    IOUtils.closeQuietly(thumbByteArrayOutputStream);
                }

                break;
            }

            case IMGSCALR: {
                pictureType = PictureType.PNG.ooxmlId;  // Workbook.PICTURE_TYPE_PNG

                final BufferedImage originalImage = ImageIO.read(url);
                if (Objects.isNull(originalImage)) {
                    throw new IOException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.UTIL_IMAGE_UNREADABLE, String.valueOf(imageFileUrl)));
                }
                final BufferedImage thumbImage = Scalr.resize(originalImage, PxlConstants.EXPORT_PICTURE_SCALE_WIDTH_IN_PIXELS, PxlConstants.EXPORT_PICTURE_SCALE_HEIGHT_IN_PIXELS);

                ByteArrayOutputStream thumbByteArrayOutputStream = null;
                try {
                    thumbByteArrayOutputStream = new ByteArrayOutputStream();

                    ImageIO.write(thumbImage, "png", thumbByteArrayOutputStream);
                    pictureBytes = thumbByteArrayOutputStream.toByteArray();
                } finally {
                    IOUtils.closeQuietly(thumbByteArrayOutputStream);
                }

                break;
            }
        }

        return Pair.of(pictureBytes, pictureType);
    }

    /**
     * Determines the POI OOXML picture-type id for an image, guessing the content type from the URL name first and then,
     * if needed, from the stream contents.
     *
     * @param url         the image source URL, used to guess the content type by name
     * @param inputStream the image stream, used to guess the content type by content when the name is inconclusive
     * @return the POI OOXML picture-type id
     * @throws IOException if the content type cannot be mapped to a supported picture type
     */
    private static int probePictureType(final String url,
                                        final InputStream inputStream)
            throws IOException {

        String contentType = URLConnection.guessContentTypeFromName(url);
        if (Objects.isNull(contentType) && Objects.nonNull(inputStream)) {
            contentType = URLConnection.guessContentTypeFromStream(inputStream);
        }

        for (final PictureType pictureType : PictureType.values()) {
            if (pictureType.contentType.equals(contentType) && pictureType.ooxmlId > 0) {
                return pictureType.ooxmlId;
            }
        }

        throw new IOException("unsupported picture type");
    }

}
