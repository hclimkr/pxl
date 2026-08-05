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

import java.time.DateTimeException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Objects;
import java.util.Optional;

/**
 * Codec for {@link LocalDate} column values — parses cells/strings into {@link LocalDate} on
 * import and writes {@link LocalDate} into cells on export.
 *
 * <p>Import reads date-formatted NUMERIC cells via POI (other numerics as Excel
 * serials) and strings via the column's cached {@link DateTimeFormatter} (falling back to the built-in
 * read formatters, then ISO-8601). A BOOLEAN cell is rejected as an unsupported cell type. Export writes
 * either a formatted string or a numeric Excel-date cell.
 */
final class PxlLocalDateCodec {

    /**
     * Prevents instantiation.
     */
    private PxlLocalDateCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses the given cell into a {@link LocalDate}. Date-formatted NUMERIC cells use POI's
     * local date-time (truncated to the date); other numerics are treated as Excel
     * date serials; STRING cells are delegated to the string parser; BOOLEAN cells are rejected as an
     * unsupported cell type; BLANK cells yield {@code null}.
     *
     * @param cell       the cell to read
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link LocalDate}, or {@code null} when blank
     * @throws PxlCellCodecException if the cell type is unsupported or the numeric value is invalid
     */
    static LocalDate parseLocalDateValue(final Cell cell,
                                         final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        LocalDate localDateValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                try {
                    localDateValue = PxlDateCellSupport.readNumericCellAsLocalDateTime(cell, "LocalDate").toLocalDate();
                } catch (NumberFormatException numberFormatException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(cell), "LocalDate"), numberFormatException);
                }
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                localDateValue = parseLocalDateValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                // final boolean booleanCellValue = cell.getBooleanCellValue();
                // localDateValue = DateUtil.getLocalDateTime(BooleanUtils.toInteger(booleanCellValue)).toLocalDate();
                // break;
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return localDateValue;
    }

    /**
     * Parses a string token into a {@link LocalDate}. The column's cached formatter is tried
     * first, then the built-in read formatters, then ISO-8601. The value is trimmed when {@code importTrim}
     * is set; a blank value yields {@code null}.
     *
     * @param s          the raw string token
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link LocalDate}, or {@code null} when blank
     * @throws PxlCellCodecException if the value matches no known date format
     */
    static LocalDate parseLocalDateValue(final String s,
                                         final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        LocalDate localDateValue = null;

        final DateTimeFormatter importDateTimeFormatter = columnMeta.getImportDateTimeFormatterCache();
        if (Objects.nonNull(importDateTimeFormatter)) {
            try {
                localDateValue = LocalDate.parse(stringValue, importDateTimeFormatter);
                return localDateValue;
            } catch (DateTimeParseException dateTimeParseException) {
                // go to next parser
            }
        }

        for (final DateTimeFormatter localDateFormatter : PxlCodecConstants.localDateReadFormatters) {
            try {
                localDateValue = LocalDate.parse(stringValue, localDateFormatter);
                return localDateValue;
            } catch (DateTimeParseException dateTimeParseException) {
                // go to next parser
            }
        }

        try {
            localDateValue = LocalDate.parse(stringValue);
            return localDateValue;
        } catch (DateTimeParseException e) {
            // go to next parser
        }

        throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "LocalDate"));
    }

    /**
     * Writes the given value as a {@link LocalDate} cell and returns the exported string. A {@link String}
     * source is parsed with the export formatter; a {@link LocalDate} source is used directly. A
     * {@code null} result blanks the cell; otherwise the cell is written as a formatted string when
     * exported to string, or as a numeric Excel-date cell when no pattern/masking applies.
     *
     * @param cell       the target cell, or {@code null} to only compute the string
     * @param object     the source value (a {@link String} or {@link LocalDate})
     * @param columnMeta the resolved export metadata for this column
     * @return the exported string, or {@code null} when blank
     * @throws PxlCellCodecException if the source is unsupported or the string is not a valid date
     */
    static String buildLocalDateCell(final Cell cell,
                                     final Object object,
                                     final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        DateTimeFormatter exportDateFormatter = columnMeta.getExportDateTimeFormatterCache();
        if (Objects.isNull(exportDateFormatter)) {
            exportDateFormatter = PxlCodecConstants.localDateWriteFormatter;
        }

        LocalDate localDateValue;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                localDateValue = null;
            } else {
                try {
                    localDateValue = LocalDate.parse(stringValue, exportDateFormatter);
                } catch (DateTimeParseException dateTimeParseException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "LocalDate"), dateTimeParseException);
                }
            }
        } else if (object instanceof LocalDate) {
            localDateValue = (LocalDate) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "LocalDate"));
        }

        if (Objects.isNull(localDateValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else {
            final String cellString = makeLocalDateExportString(localDateValue, columnMeta);
            if (Objects.nonNull(cell)) {
                if (columnMeta.isExportedToString()) {
                    cell.setCellValue(cellString);
                } else {
                    // When there is no pattern/masking, write it as a Numeric (Excel date serial) cell rather than a string.
                    PxlDateCellSupport.writeNumericCell(cell, localDateValue, columnMeta, PxlCodecConstants.localDateExcelFormat);
                }
            }

            return cellString;
        }
    }

    /**
     * Renders the export string for a {@link LocalDate}: formats it with the configured export {@link DateTimeFormatter}
     * (or the built-in default write formatter when none is set), then applies string-level export processing via
     * {@link PxlStringCodec#makeExportString}.
     *
     * @param localDateValue the value to render
     * @param columnMeta     resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the value cannot be formatted with the export formatter
     */
    private static String makeLocalDateExportString(final LocalDate localDateValue,
                                                    final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(localDateValue)) {
            return null;
        }

        DateTimeFormatter exportDateFormatter = columnMeta.getExportDateTimeFormatterCache();
        if (Objects.isNull(exportDateFormatter)) {
            exportDateFormatter = PxlCodecConstants.localDateWriteFormatter;
        }

        try {
            final String stringValue = localDateValue.format(exportDateFormatter);

            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } catch (DateTimeException dateTimeException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(localDateValue), "LocalDate"), dateTimeException);
        }
    }

}
