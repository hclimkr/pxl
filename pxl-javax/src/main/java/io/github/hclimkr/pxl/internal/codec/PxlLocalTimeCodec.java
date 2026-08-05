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
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Objects;
import java.util.Optional;

/**
 * Codec for {@link LocalTime} column values — parses cells/strings into {@link LocalTime} on
 * import and writes {@link LocalTime} into cells on export.
 *
 * <p>Import reads NUMERIC cells as Excel time fractions and strings via the column's cached
 * {@link DateTimeFormatter} (falling back to the built-in read formatters, then ISO-8601). A BOOLEAN cell
 * is rejected as an unsupported cell type. Export writes either a formatted string or a numeric Excel-time
 * cell.
 */
final class PxlLocalTimeCodec {

    /**
     * Prevents instantiation.
     */
    private PxlLocalTimeCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses the given cell into a {@link LocalTime}. NUMERIC cells are read as Excel time
     * fractions (via POI's local date-time truncated to the time); STRING cells are delegated to the string
     * parser; BOOLEAN cells are rejected as an unsupported cell type; BLANK cells yield {@code null}.
     *
     * @param cell       the cell to read
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link LocalTime}, or {@code null} when blank
     * @throws PxlCellCodecException if the cell type is unsupported or the numeric value is invalid
     */
    static LocalTime parseLocalTimeValue(final Cell cell,
                                         final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        LocalTime localTimeValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                try {
                    localTimeValue = PxlDateCellSupport.readNumericCellAsLocalDateTime(cell, "LocalTime").toLocalTime();
                } catch (NumberFormatException numberFormatException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(cell), "LocalTime"), numberFormatException);
                }
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                localTimeValue = parseLocalTimeValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                // final boolean booleanCellValue = cell.getBooleanCellValue();
                // localTimeValue = DateUtil.getLocalDateTime(BooleanUtils.toInteger(booleanCellValue)).toLocalTime();
                // break;
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return localTimeValue;
    }

    /**
     * Parses a string token into a {@link LocalTime}. The column's cached formatter is tried
     * first, then the built-in read formatters, then ISO-8601. The value is trimmed when {@code importTrim}
     * is set; a blank value yields {@code null}.
     *
     * @param s          the raw string token
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link LocalTime}, or {@code null} when blank
     * @throws PxlCellCodecException if the value matches no known time format
     */
    static LocalTime parseLocalTimeValue(final String s,
                                         final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        LocalTime localTimeValue = null;

        final DateTimeFormatter importDateTimeFormatter = columnMeta.getImportDateTimeFormatterCache();
        if (Objects.nonNull(importDateTimeFormatter)) {
            try {
                localTimeValue = LocalTime.parse(stringValue, importDateTimeFormatter);
                return localTimeValue;
            } catch (DateTimeParseException dateTimeParseException) {
                // go to next parser
            }
        }

        for (final DateTimeFormatter localTimeFormatter : PxlCodecConstants.localTimeReadFormatters) {
            try {
                localTimeValue = LocalTime.parse(stringValue, localTimeFormatter);
                return localTimeValue;
            } catch (DateTimeParseException dateTimeParseException) {
                // go to next parser
            }
        }

        try {
            localTimeValue = LocalTime.parse(stringValue);
            return localTimeValue;
        } catch (DateTimeParseException e) {
            // go to next parser
        }

        throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "LocalTime"));
    }

    /**
     * Writes the given value as a {@link LocalTime} cell and returns the exported string. A {@link String}
     * source is parsed with the export formatter; a {@link LocalTime} source is used directly. A
     * {@code null} result blanks the cell; otherwise the cell is written as a formatted string when
     * exported to string, or as a numeric Excel-time cell when no pattern/masking applies.
     *
     * @param cell       the target cell, or {@code null} to only compute the string
     * @param object     the source value (a {@link String} or {@link LocalTime})
     * @param columnMeta the resolved export metadata for this column
     * @return the exported string, or {@code null} when blank
     * @throws PxlCellCodecException if the source is unsupported or the string is not a valid time
     */
    static String buildLocalTimeCell(final Cell cell,
                                     final Object object,
                                     final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        DateTimeFormatter exportTimeFormatter = columnMeta.getExportDateTimeFormatterCache();
        if (Objects.isNull(exportTimeFormatter)) {
            exportTimeFormatter = PxlCodecConstants.localTimeWriteFormatter;
        }

        LocalTime localTimeValue;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                localTimeValue = null;
            } else {
                try {
                    localTimeValue = LocalTime.parse(stringValue, exportTimeFormatter);
                } catch (DateTimeParseException dateTimeParseException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "LocalTime"), dateTimeParseException);
                }
            }
        } else if (object instanceof LocalTime) {
            localTimeValue = (LocalTime) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "LocalTime"));
        }

        if (Objects.isNull(localTimeValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else {
            final String cellString = makeLocalTimeExportString(localTimeValue, columnMeta);
            if (Objects.nonNull(cell)) {
                if (columnMeta.isExportedToString()) {
                    cell.setCellValue(cellString);
                } else {
                    // When there is no pattern/masking, write it as a Numeric (Excel time fraction) cell rather than a string.
                    PxlDateCellSupport.writeNumericTimeCell(cell, localTimeValue, columnMeta, PxlCodecConstants.localTimeExcelFormat);
                }
            }

            return cellString;
        }
    }

    /**
     * Renders the export string for a {@link LocalTime}: formats it with the configured export {@link DateTimeFormatter}
     * (or the built-in default write formatter when none is set), then applies string-level export processing via
     * {@link PxlStringCodec#makeExportString}.
     *
     * @param localTimeValue the value to render
     * @param columnMeta     resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the value cannot be formatted with the export formatter
     */
    private static String makeLocalTimeExportString(final LocalTime localTimeValue,
                                                    final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(localTimeValue)) {
            return null;
        }

        DateTimeFormatter exportTimeFormatter = columnMeta.getExportDateTimeFormatterCache();
        if (Objects.isNull(exportTimeFormatter)) {
            exportTimeFormatter = PxlCodecConstants.localTimeWriteFormatter;
        }

        try {
            final String stringValue = localTimeValue.format(exportTimeFormatter);

            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } catch (DateTimeException dateTimeException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(localTimeValue), "LocalTime"), dateTimeException);
        }
    }

}
