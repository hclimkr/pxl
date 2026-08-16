package io.github.hclimkr.pxl.internal.codec;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlExportColumnMeta;
import io.github.hclimkr.pxl.internal.meta.PxlImportColumnMeta;
import io.github.hclimkr.pxl.util.PxlCellUtils;
import org.apache.commons.lang3.BooleanUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;

import java.util.Arrays;
import java.util.Objects;
import java.util.regex.Pattern;

/**
 * Codec for {@link String} column values - reads a cell (or CSV token) into a {@link String} on import
 * and writes a {@link String} into a cell on export.
 *
 * <p>On import, NUMERIC cells are rendered via the workbook's cached {@link DataFormatter} (built with
 * {@code Locale.ROOT} so decimal/grouping symbols are locale-independent; streaming cells included - the
 * streaming reader reads styles by default, so the cell carries its number format), BOOLEAN cells via the
 * column's import true/false strings, and BLANK cells map to {@code null}. On export, a leading
 * {@code '='} is written as a formula (or a quote-prefixed literal), pictures are embedded when requested,
 * and trim/masking options are applied.
 */
final class PxlStringCodec {

    /**
     * Prevents instantiation.
     */
    private PxlStringCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses the given cell into a {@link String}. NUMERIC cells are formatted with the workbook's cached
     * {@link DataFormatter} (streaming cells included - the streaming reader reads styles by default, so the
     * cell carries its number format); STRING cells are normalized per the column's {@code importTrim}
     * option; BOOLEAN cells are rendered with the column's import true/false strings; BLANK cells yield
     * {@code null}.
     *
     * @param cell       the cell to read
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link String}, or {@code null} when the cell is blank
     * @throws PxlCellCodecException if the cell type is not supported
     */
    static String parseStringValue(final Cell cell,
                                   final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        String stringValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                // #1 With this approach, the result rendered by applying the cell's display format and locale is turned into text
                final DataFormatter dataFormatter = columnMeta.getWorkbookMeta().getImportDataFormatterCache();
                stringValue = dataFormatter.formatCellValue(cell);

//                // #2 With this approach the double held in the cell is turned into text as-is, ignoring the cell format
//                final double numericCellValue = cell.getNumericCellValue();
//                stringValue = NumberToTextConverter.toText(numericCellValue);

//                // #3: With this approach there is an issue where a cell value of 2012000046 is converted to the string "2.012000046E9". (#1215)
//                final double numericCellValue = cell.getNumericCellValue();
//                stringValue = String.valueOf(numericCellValue);

                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                stringValue = parseStringValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                final boolean booleanCellValue = cell.getBooleanCellValue();
                final String importTrueString = columnMeta.getImportTrueString();
                final String importFalseString = columnMeta.getImportFalseString();
                stringValue = BooleanUtils.toString(booleanCellValue, importTrueString, importFalseString);
                break;

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return stringValue;
    }

    /**
     * Normalizes a raw string token (CSV path): trims it when {@code importTrim} is set and maps an empty
     * result to {@code null}.
     *
     * @param s          the raw string token
     * @param columnMeta the resolved import metadata for this column
     * @return the normalized value, or {@code null} when empty
     */
    static String parseStringValue(final String s,
                                   final PxlImportColumnMeta columnMeta) {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;

        return StringUtils.isEmpty(stringValue) ? null : stringValue;
    }

    /**
     * Writes the given value's {@link String} form into the cell (when non-{@code null}) and returns it.
     *
     * <p>The column's options pick the form, in this order: with {@code exportStringAsFormula} enabled a value
     * starting with {@code '='} is set as a cell formula (falling back to a quote-prefixed literal when POI rejects
     * it), then {@code exportStringAsPicture} renders the string as an embedded picture, and otherwise the string is
     * written as a plain value. Formula therefore wins when both options are set. A value starting with {@code '='}
     * that lands in the plain-text form is written quote-prefixed, so Excel shows it verbatim instead of reading it
     * as a formula - a safeguard on how text is written, not a fourth way of choosing the form.
     *
     * @param cell       the target cell, or {@code null} to only compute the string
     * @param object     the source value (a {@link String})
     * @param columnMeta the resolved export metadata for this column
     * @return the exported string
     */
    static String buildStringCell(final Cell cell,
                                  final Object object,
                                  final PxlExportColumnMeta columnMeta) {

        final String cellString = makeExportString(object, columnMeta);
        if (Objects.nonNull(cell)) {
            if (columnMeta.isExportStringAsFormula() && StringUtils.startsWith(cellString, "=")) {  // FORMULA
                try {
                    cell.setCellFormula(StringUtils.substring(cellString, 1));
                } catch (Exception ignored) {
                    columnMeta.setQuotePrefixedCellValue(cell, cellString);
                }
            } else if (columnMeta.isExportStringAsPicture()) {
                PxlCellUtils.addPicturesToCell(cell,
                        Arrays.asList(cellString),
                        PxlConstants.EXPORT_PICTURE_SCREEN_WIDTH_IN_PIXELS,
                        PxlConstants.EXPORT_PICTURE_SCREEN_HEIGHT_IN_PIXELS,
                        PxlConstants.EXPORT_PICTURE_SCREEN_PADDING_IN_PIXELS,
                        PxlConstants.EXPORT_HORIZONTAL_NUMBER_OF_PICTURE);
            } else if (StringUtils.startsWith(cellString, "=")) {
                // Quote(Apostrophe) Prefix
                // A leading single quote character ' is a special character in Excel, which tells it to treat it's contents verbatim.
                columnMeta.setQuotePrefixedCellValue(cell, cellString);
            } else {
                cell.setCellValue(cellString);
            }
        }

        return cellString;
    }

    /**
     * Applies the column's export trim and masking options to a raw string value. Shared by the other
     * codecs to post-process their string representation.
     *
     * @param object     the source value (a {@link String})
     * @param columnMeta the resolved export metadata for this column
     * @return the trimmed and/or masked string, or {@code null} when the input is {@code null}
     */
    static String makeExportString(final Object object,
                                   final PxlExportColumnMeta columnMeta) {

        final String stringValue = columnMeta.isExportTrim() ? StringUtils.trim((String) object) : (String) object;
        final Pattern exportMaskingPattern = columnMeta.getExportMaskingPattern();

        if (Objects.nonNull(exportMaskingPattern) && Objects.nonNull(stringValue)) {
            return exportMaskingPattern.matcher(stringValue).replaceAll(PxlConstants.DEFAULT_EXPORT_MASKING_CHAR);
        } else {
            return stringValue;
        }
    }

}
