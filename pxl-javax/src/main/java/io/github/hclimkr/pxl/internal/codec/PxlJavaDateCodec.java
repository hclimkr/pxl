package io.github.hclimkr.pxl.internal.codec;

import io.github.hclimkr.pxl.exception.PxlCellCodecException;
import io.github.hclimkr.pxl.internal.constant.PxlCodecConstants;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlExportColumnMeta;
import io.github.hclimkr.pxl.internal.meta.PxlImportColumnMeta;
import io.github.hclimkr.pxl.internal.support.PxlDateCellSupport;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.format.DateTimeParseException;
import java.util.Date;
import java.util.Objects;
import java.util.Optional;

/**
 * Codec for {@link Date} column values — parses cells/strings into {@link Date} on import and
 * writes {@link Date} into cells on export.
 *
 * <p>Import reads date-formatted NUMERIC cells via POI, other numerics as Excel date
 * serials, and strings via the column's cached {@link SimpleDateFormat} (falling back to the built-in read
 * formatters, then an ISO-8601 instant). A BOOLEAN cell is rejected as an unsupported cell type. Export
 * writes either a formatted string (when exported to string) or a numeric Excel-date cell (when no
 * pattern/masking applies).
 */
final class PxlJavaDateCodec {

    /**
     * Prevents instantiation.
     */
    private PxlJavaDateCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses the given cell into a {@link Date}. Date-formatted NUMERIC cells use POI's date
     * value; other numerics are treated as Excel date serials via
     * {@link DateUtil#getJavaDate}; STRING cells are delegated to the string parser; BOOLEAN cells are
     * rejected as an unsupported cell type; BLANK cells yield {@code null}.
     *
     * @param cell       the cell to read
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link Date}, or {@code null} when blank
     * @throws PxlCellCodecException if the cell type is unsupported or the numeric value is invalid
     */
    static Date parseJavaDateValue(final Cell cell,
                                   final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        Date dateValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                try {
                    dateValue = PxlDateCellSupport.readNumericCellAsJavaDate(cell, "Date");
                } catch (NumberFormatException numberFormatException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(cell), "Date"), numberFormatException);
                }
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                dateValue = parseJavaDateValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                // final boolean booleanCellValue = cell.getBooleanCellValue();
                // dateValue = DateUtil.getJavaDate(BooleanUtils.toInteger(booleanCellValue));
                // break;
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return dateValue;
    }

    /**
     * Parses a string token into a {@link Date}. The column's cached {@link SimpleDateFormat} is
     * tried first, then the built-in read formatters, then an ISO-8601 instant. The value is trimmed when
     * {@code importTrim} is set; a blank value yields {@code null}.
     *
     * @param s          the raw string token
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link Date}, or {@code null} when blank
     * @throws PxlCellCodecException if the value matches no known date format
     */
    static Date parseJavaDateValue(final String s,
                                   final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        Date dateValue = null;

        final SimpleDateFormat importJavaDateFormatter = columnMeta.getImportJavaDateFormatterCache();
        if (Objects.nonNull(importJavaDateFormatter)) {
            try {
                dateValue = importJavaDateFormatter.parse(stringValue);
                return dateValue;
            } catch (ParseException parseException) {
                // go to next parser
            }
        }

        for (final SimpleDateFormat simpleDateFormatter : PxlCodecConstants.javaDateReadFormatters.get()) {
            try {
                dateValue = simpleDateFormatter.parse(stringValue);
                return dateValue;
            } catch (ParseException parseException) {
                // go to next parser
            }
        }

        try {
            dateValue = Date.from(Instant.parse(stringValue));
            return dateValue;
        } catch (DateTimeParseException | IllegalArgumentException e) {
            // go to next parser
        }

        throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "Date"));
    }

    /**
     * Writes the given value as a {@link Date} cell and returns the exported string. A {@link String}
     * source is parsed with the export formatter; a {@link Date} source is used directly. A
     * {@code null} result blanks the cell; otherwise the cell is written as a formatted string when
     * exported to string, or as a numeric Excel-date cell when no pattern/masking applies.
     *
     * @param cell       the target cell, or {@code null} to only compute the string
     * @param object     the source value (a {@link String} or {@link Date})
     * @param columnMeta the resolved export metadata for this column
     * @return the exported string, or {@code null} when blank
     * @throws PxlCellCodecException if the source is unsupported or the string is not a valid date
     */
    static String buildJavaDateCell(final Cell cell,
                                    final Object object,
                                    final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        SimpleDateFormat exportJavaDateFormatter = columnMeta.getExportJavaDateFormatterCache();
        if (Objects.isNull(exportJavaDateFormatter)) {
            exportJavaDateFormatter = PxlCodecConstants.javaDateWriteFormatter.get();
        }

        Date dateValue;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                dateValue = null;
            } else {
                try {
                    dateValue = exportJavaDateFormatter.parse(stringValue);
                } catch (ParseException parseException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "Date"), parseException);
                }
            }
        } else if (object instanceof Date) {
            dateValue = (Date) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "Date"));
        }

        if (Objects.isNull(dateValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else {
            final String cellString = makeJavaDateExportString(dateValue, columnMeta);
            if (Objects.nonNull(cell)) {
                if (columnMeta.isExportedToString()) {
                    cell.setCellValue(cellString);
                } else {
                    // When there is no pattern/masking, write it as a Numeric (Excel date serial) cell rather than a string.
                    PxlDateCellSupport.writeNumericCell(cell, dateValue, columnMeta, PxlCodecConstants.javaDateExcelFormat);
                }
            }

            return cellString;
        }
    }

    /**
     * Renders the export string for a {@link Date}: formats it with the configured export {@link SimpleDateFormat}
     * (or the built-in default write formatter when none is set), then applies string-level export processing via
     * {@link PxlStringCodec#makeExportString}.
     *
     * @param dateValue  the value to render
     * @param columnMeta resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the export pattern cannot be applied to the value
     */
    private static String makeJavaDateExportString(final Date dateValue,
                                                   final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(dateValue)) {
            return null;
        }

        SimpleDateFormat exportJavaDateFormatter = columnMeta.getExportJavaDateFormatterCache();
        if (Objects.isNull(exportJavaDateFormatter)) {
            exportJavaDateFormatter = PxlCodecConstants.javaDateWriteFormatter.get();
        }

        final String stringValue = exportJavaDateFormatter.format(dateValue);

        return PxlStringCodec.makeExportString(stringValue, columnMeta);
    }

}
