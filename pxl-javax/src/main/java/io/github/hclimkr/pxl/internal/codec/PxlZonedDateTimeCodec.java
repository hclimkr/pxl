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
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Objects;
import java.util.Optional;

/**
 * Codec for {@link ZonedDateTime} column values - writes {@link ZonedDateTime} into cells on export
 * and parses cells/strings into {@link ZonedDateTime} on import.
 *
 * <p>On export to a numeric cell only the local (wall-clock) part is stored, since Excel serials carry no
 * zone. On import, NUMERIC cells are read as Excel serials and placed at the system default zone, a BOOLEAN
 * cell is rejected as an unsupported cell type, and strings are parsed with the column's cached formatter,
 * then ISO-8601 with an explicit offset/zone ({@link DateTimeFormatter#ISO_ZONED_DATE_TIME}).
 */
final class PxlZonedDateTimeCodec {

    /**
     * Prevents instantiation.
     */
    private PxlZonedDateTimeCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Writes the given value as a {@link ZonedDateTime} cell and returns the exported string. A
     * {@link String} source is parsed with the export formatter, falling back to ISO-8601; a
     * {@link ZonedDateTime} source is used directly. A {@code null} result blanks the cell;
     * otherwise the cell is written as a formatted string when exported to string, or as a numeric
     * Excel-date cell (local part only) when no pattern/masking applies.
     *
     * @param cell       the target cell, or {@code null} to only compute the string
     * @param object     the source value (a {@link String} or {@link ZonedDateTime})
     * @param columnMeta the resolved export metadata for this column
     * @return the exported string, or {@code null} when blank
     * @throws PxlCellCodecException if the source is unsupported or the string is not a valid zoned date-time
     */
    static String buildZonedDateTimeCell(final Cell cell,
                                         final Object object,
                                         final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final DateTimeFormatter exportDateTimeFormatter = columnMeta.getExportDateTimeFormatterCache();
        ZonedDateTime zonedDateTimeValue = null;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                zonedDateTimeValue = null;
            } else {
                if (Objects.nonNull(exportDateTimeFormatter)) {
                    try {
                        zonedDateTimeValue = ZonedDateTime.parse(stringValue, exportDateTimeFormatter);
                    } catch (DateTimeParseException dateTimeParseException) {
                        // go to next parser
                    }
                }

                if (Objects.isNull(zonedDateTimeValue)) {
                    try {
                        zonedDateTimeValue = ZonedDateTime.parse(stringValue);
                    } catch (DateTimeParseException isoParseException) {
                        throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(stringValue), "ZonedDateTime"), isoParseException);
                    }
                }
            }
        } else if (object instanceof ZonedDateTime) {
            zonedDateTimeValue = (ZonedDateTime) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "ZonedDateTime"));
        }

        if (Objects.isNull(zonedDateTimeValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else {
            final String cellString = makeZonedDateTimeExportString(zonedDateTimeValue, columnMeta);
            if (Objects.nonNull(cell)) {
                if (columnMeta.isExportedToString()) {
                    cell.setCellValue(cellString);
                } else {
                    // When there is no pattern/masking, write it as a Numeric (Excel date serial) cell.
                    // Excel serials have no notion of zone/offset, so only the wall-clock (local) part is stored (the same loss as string export).
                    PxlDateCellSupport.writeNumericCell(cell, zonedDateTimeValue.toLocalDateTime(), columnMeta, PxlCodecConstants.localDateTimeExcelFormat);
                }
            }

            return cellString;
        }
    }

    /**
     * Renders the export string for a {@link ZonedDateTime}: formats it with the configured export {@link DateTimeFormatter}
     * (or {@link DateTimeFormatter#ISO_ZONED_DATE_TIME} when none is set), then applies string-level export processing via
     * {@link PxlStringCodec#makeExportString}.
     *
     * @param zonedDateTimeValue the value to render
     * @param columnMeta         resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the value cannot be formatted with the export formatter
     */
    private static String makeZonedDateTimeExportString(final ZonedDateTime zonedDateTimeValue,
                                                        final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(zonedDateTimeValue)) {
            return null;
        }

        final DateTimeFormatter exportDateTimeFormatter = columnMeta.getExportDateTimeFormatterCache();
        try {
            final String stringValue = Objects.nonNull(exportDateTimeFormatter)
                    ? zonedDateTimeValue.format(exportDateTimeFormatter)
                    : zonedDateTimeValue.format(DateTimeFormatter.ISO_ZONED_DATE_TIME);

            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } catch (DateTimeException dateTimeException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(zonedDateTimeValue), "ZonedDateTime"), dateTimeException);
        }
    }

    /**
     * Parses the given cell into a {@link ZonedDateTime} at the system default zone.
     * Date-formatted NUMERIC cells use POI's local date-time; other numerics are
     * Excel serials; STRING cells are delegated to the string parser; BOOLEAN cells are rejected as an
     * unsupported cell type; BLANK cells yield {@code null}.
     *
     * @param cell       the cell to read
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link ZonedDateTime}, or {@code null} when blank
     * @throws PxlCellCodecException if the cell type is unsupported or the numeric value is invalid
     */
    static ZonedDateTime parseZonedDateTimeValue(final Cell cell,
                                                 final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final ZoneId zoneId = ZoneId.systemDefault();
        ZonedDateTime zonedDateTimeValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                try {
                    zonedDateTimeValue = PxlDateCellSupport.readNumericCellAsLocalDateTime(cell, "ZonedDateTime").atZone(zoneId);
                } catch (NumberFormatException numberFormatException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(cell), "ZonedDateTime"), numberFormatException);
                }
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                zonedDateTimeValue = parseZonedDateTimeValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                // final boolean booleanCellValue = cell.getBooleanCellValue();
                // zonedDateTimeValue = DateUtil.getLocalDateTime(BooleanUtils.toInteger(booleanCellValue)).atZone(zoneId);
                // break;
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return zonedDateTimeValue;
    }

    /**
     * Parses a string token into a {@link ZonedDateTime}. The column's cached formatter is tried
     * first, then ISO-8601 parsing with an explicit offset/zone. The value is trimmed when {@code importTrim}
     * is set; a blank value yields {@code null}.
     *
     * @param s          the raw string token
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link ZonedDateTime}, or {@code null} when blank
     * @throws PxlCellCodecException if the value matches no known format
     */
    static ZonedDateTime parseZonedDateTimeValue(final String s,
                                                 final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        ZonedDateTime zonedDateTimeValue = null;

        final DateTimeFormatter importDateTimeFormatter = columnMeta.getImportDateTimeFormatterCache();
        if (Objects.nonNull(importDateTimeFormatter)) {
            try {
                zonedDateTimeValue = ZonedDateTime.parse(stringValue, importDateTimeFormatter);
                return zonedDateTimeValue;
            } catch (DateTimeParseException dateTimeParseException) {
                // go to next parser
            }
        }

        try {
            zonedDateTimeValue = ZonedDateTime.parse(stringValue, DateTimeFormatter.ISO_ZONED_DATE_TIME);
            return zonedDateTimeValue;
        } catch (DateTimeParseException e) {
            // go to next parser
        }

        throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "ZonedDateTime"));
    }

}
