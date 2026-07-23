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

import java.time.DateTimeException;
import java.time.OffsetDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Objects;
import java.util.Optional;

/**
 * Codec for {@link java.time.OffsetDateTime} column values — parses cells/strings into {@code OffsetDateTime}
 * on import and writes {@code OffsetDateTime} into cells on export.
 *
 * <p>NUMERIC cells are read as Excel serials and given the system zone's offset for that date; a BOOLEAN
 * cell is rejected as an unsupported cell type. Strings are parsed with the column's cached formatter, then
 * ISO-8601 with an explicit offset ({@link DateTimeFormatter#ISO_OFFSET_DATE_TIME}). On export to a numeric
 * cell only the local (wall-clock) part is stored, since Excel serials carry no offset.
 */
final class PxlOffsetDateTimeCodec {

    /**
     * Prevents instantiation.
     */
    private PxlOffsetDateTimeCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses the given cell into an {@link java.time.OffsetDateTime} using the system default zone's offset.
     * Date-formatted NUMERIC cells use POI's local date-time; other numerics are
     * Excel serials; STRING cells are delegated to the string parser; BOOLEAN cells are rejected as an
     * unsupported cell type; BLANK cells yield {@code null}.
     *
     * @param cell       the cell to read
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@code OffsetDateTime}, or {@code null} when blank
     * @throws PxlCellCodecException if the cell type is unsupported or the numeric value is invalid
     */
    static OffsetDateTime parseOffsetDateTimeValue(final Cell cell,
                                                   final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        OffsetDateTime offsetDateTimeValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                final double numericValue = cell.getNumericCellValue();
                try {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        offsetDateTimeValue = cell.getLocalDateTimeCellValue().atZone(ZoneId.systemDefault()).toOffsetDateTime();
                    } else {
                        offsetDateTimeValue = DateUtil.getLocalDateTime(numericValue).atZone(ZoneId.systemDefault()).toOffsetDateTime();
                    }
                } catch (NumberFormatException numberFormatException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(cell), "OffsetDateTime"), numberFormatException);
                }
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                offsetDateTimeValue = parseOffsetDateTimeValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                // final boolean booleanCellValue = cell.getBooleanCellValue();
                // offsetDateTimeValue = DateUtil.getLocalDateTime(BooleanUtils.toInteger(booleanCellValue)).atZone(ZoneId.systemDefault()).toOffsetDateTime();
                // break;
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return offsetDateTimeValue;
    }

    /**
     * Parses a string token into an {@link java.time.OffsetDateTime}. The column's cached formatter is
     * tried first, then ISO-8601 parsing with an explicit offset. The value is trimmed when {@code importTrim}
     * is set; a blank value yields {@code null}.
     *
     * @param s          the raw string token
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@code OffsetDateTime}, or {@code null} when blank
     * @throws PxlCellCodecException if the value matches no known format
     */
    static OffsetDateTime parseOffsetDateTimeValue(final String s,
                                                   final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        OffsetDateTime offsetDateTimeValue = null;

        final DateTimeFormatter importDateTimeFormatter = columnMeta.getImportDateTimeFormatterCache();
        if (Objects.nonNull(importDateTimeFormatter)) {
            try {
                offsetDateTimeValue = OffsetDateTime.parse(stringValue, importDateTimeFormatter);
                return offsetDateTimeValue;
            } catch (DateTimeParseException dateTimeParseException) {
                // go to next parser
            }
        }

        try {
            offsetDateTimeValue = OffsetDateTime.parse(stringValue, DateTimeFormatter.ISO_OFFSET_DATE_TIME);
            return offsetDateTimeValue;
        } catch (DateTimeParseException e) {
            // go to next parser
        }

        throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "OffsetDateTime"));
    }

    /**
     * Writes the given value as an {@code OffsetDateTime} cell and returns the exported string. A
     * {@code String} source is parsed with the export formatter, falling back to ISO-8601; a
     * {@link java.time.OffsetDateTime} source is used directly. A {@code null} result blanks the cell;
     * otherwise the cell is written as a formatted string when exported to string, or as a numeric
     * Excel-date cell (local part only) when no pattern/masking applies.
     *
     * @param cell       the target cell, or {@code null} to only compute the string
     * @param object     the source value (a {@code String} or {@code OffsetDateTime})
     * @param columnMeta the resolved export metadata for this column
     * @return the exported string, or {@code null} when blank
     * @throws PxlCellCodecException if the source is unsupported or the string is not a valid offset date-time
     */
    static String buildOffsetDateTimeCell(final Cell cell,
                                          final Object object,
                                          final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final DateTimeFormatter exportDateTimeFormatter = columnMeta.getExportDateTimeFormatterCache();
        OffsetDateTime offsetDateTimeValue = null;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                offsetDateTimeValue = null;
            } else {
                if (Objects.nonNull(exportDateTimeFormatter)) {
                    try {
                        offsetDateTimeValue = OffsetDateTime.parse(stringValue, exportDateTimeFormatter);
                    } catch (DateTimeParseException dateTimeParseException) {
                        // go to next parser
                    }
                }

                if (Objects.isNull(offsetDateTimeValue)) {
                    try {
                        offsetDateTimeValue = OffsetDateTime.parse(stringValue);
                    } catch (DateTimeParseException isoParseException) {
                        throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "OffsetDateTime"), isoParseException);
                    }
                }
            }
        } else if (object instanceof OffsetDateTime) {
            offsetDateTimeValue = (OffsetDateTime) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "OffsetDateTime"));
        }

        if (Objects.isNull(offsetDateTimeValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else {
            final String cellString = makeOffsetDateTimeExportString(offsetDateTimeValue, columnMeta);
            if (Objects.nonNull(cell)) {
                if (columnMeta.isExportedToString()) {
                    cell.setCellValue(cellString);
                } else {
                    // When there is no pattern/masking, write it as a Numeric (Excel date serial) cell.
                    // Excel serials have no notion of offset, so only the wall-clock (local) part is stored (the same loss as string export).
                    PxlDateCellSupport.writeNumericCell(cell, offsetDateTimeValue.toLocalDateTime(), columnMeta, PxlCodecConstants.localDateTimeExcelFormat);
                }
            }

            return cellString;
        }
    }

    /**
     * Renders the export string for an {@code OffsetDateTime}: formats it with the configured export {@link DateTimeFormatter}
     * (or {@link DateTimeFormatter#ISO_OFFSET_DATE_TIME} when none is set), then applies string-level export processing via
     * {@link PxlStringCodec#makeExportString}.
     *
     * @param offsetDateTimeValue the value to render
     * @param columnMeta          resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the value cannot be formatted with the export formatter
     */
    private static String makeOffsetDateTimeExportString(final OffsetDateTime offsetDateTimeValue,
                                                         final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(offsetDateTimeValue)) {
            return null;
        }

        final DateTimeFormatter exportDateTimeFormatter = columnMeta.getExportDateTimeFormatterCache();
        try {
            final String stringValue = Objects.nonNull(exportDateTimeFormatter)
                    ? offsetDateTimeValue.format(exportDateTimeFormatter)
                    : offsetDateTimeValue.format(DateTimeFormatter.ISO_OFFSET_DATE_TIME);

            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } catch (DateTimeException dateTimeException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(offsetDateTimeValue), "OffsetDateTime"), dateTimeException);
        }
    }

}
