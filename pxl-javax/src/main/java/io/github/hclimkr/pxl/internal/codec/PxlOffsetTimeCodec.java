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
import java.time.OffsetTime;
import java.time.ZoneId;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Objects;
import java.util.Optional;

/**
 * Codec for {@link OffsetTime} column values - parses cells/strings into {@link OffsetTime} on
 * import and writes {@link OffsetTime} into cells on export.
 *
 * <p>NUMERIC cells are read as Excel time fractions and given the current system offset; a BOOLEAN cell is
 * rejected as an unsupported cell type. Strings are parsed with the column's cached formatter, then ISO-8601
 * with an explicit offset ({@link DateTimeFormatter#ISO_OFFSET_TIME}). On export to a numeric cell only the
 * local (wall-clock) time is stored.
 */
final class PxlOffsetTimeCodec {

    /**
     * Prevents instantiation.
     */
    private PxlOffsetTimeCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses the given cell into an {@link OffsetTime} using the current system offset. NUMERIC
     * cells are read as Excel time fractions; STRING cells are delegated to the string parser; BOOLEAN
     * cells are rejected as an unsupported cell type; BLANK cells yield {@code null}.
     *
     * @param cell       the cell to read
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link OffsetTime}, or {@code null} when blank
     * @throws PxlCellCodecException if the cell type is unsupported or the numeric value is invalid
     */
    static OffsetTime parseOffsetTimeValue(final Cell cell,
                                           final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final ZoneOffset zoneOffset = OffsetTime.now(ZoneId.systemDefault()).getOffset();
        OffsetTime offsetTimeValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                try {
                    offsetTimeValue = PxlDateCellSupport.readNumericCellAsLocalDateTime(cell, "OffsetTime").toLocalTime().atOffset(zoneOffset);
                } catch (NumberFormatException numberFormatException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(cell), "OffsetTime"), numberFormatException);
                }
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                offsetTimeValue = parseOffsetTimeValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                // final boolean booleanCellValue = cell.getBooleanCellValue();
                // offsetTimeValue = DateUtil.getLocalDateTime(BooleanUtils.toInteger(booleanCellValue)).toLocalTime().atOffset(zoneOffset);
                // break;
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return offsetTimeValue;
    }

    /**
     * Parses a string token into an {@link OffsetTime}. The column's cached formatter is tried
     * first, then ISO-8601 parsing with an explicit offset. The value is trimmed when {@code importTrim} is
     * set; a blank value yields {@code null}.
     *
     * @param s          the raw string token
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@link OffsetTime}, or {@code null} when blank
     * @throws PxlCellCodecException if the value matches no known format
     */
    static OffsetTime parseOffsetTimeValue(final String s,
                                           final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        OffsetTime offsetTimeValue = null;

        final DateTimeFormatter importDateTimeFormatter = columnMeta.getImportDateTimeFormatterCache();
        if (Objects.nonNull(importDateTimeFormatter)) {
            try {
                offsetTimeValue = OffsetTime.parse(stringValue, importDateTimeFormatter);
                return offsetTimeValue;
            } catch (DateTimeParseException dateTimeParseException) {
                // go to next parser
            }
        }

        try {
            offsetTimeValue = OffsetTime.parse(stringValue, DateTimeFormatter.ISO_OFFSET_TIME);
            return offsetTimeValue;
        } catch (DateTimeParseException e) {
            // go to next parser
        }

        throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_IMPORT_PARSE_INVALID, String.valueOf(stringValue), "OffsetTime"));
    }

    /**
     * Writes the given value as an {@link OffsetTime} cell and returns the exported string. A {@link String}
     * source is parsed with the export formatter, falling back to ISO-8601; a {@link OffsetTime}
     * source is used directly. A {@code null} result blanks the cell; otherwise the cell is written as a
     * formatted string when exported to string, or as a numeric Excel-time cell (local part only) when no
     * pattern/masking applies.
     *
     * @param cell       the target cell, or {@code null} to only compute the string
     * @param object     the source value (a {@link String} or {@link OffsetTime})
     * @param columnMeta the resolved export metadata for this column
     * @return the exported string, or {@code null} when blank
     * @throws PxlCellCodecException if the source is unsupported or the string is not a valid offset time
     */
    static String buildOffsetTimeCell(final Cell cell,
                                      final Object object,
                                      final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final DateTimeFormatter exportTimeFormatter = columnMeta.getExportDateTimeFormatterCache();
        OffsetTime offsetTimeValue = null;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                offsetTimeValue = null;
            } else {
                if (Objects.nonNull(exportTimeFormatter)) {
                    try {
                        offsetTimeValue = OffsetTime.parse(stringValue, exportTimeFormatter);
                    } catch (DateTimeParseException dateTimeParseException) {
                        // go to next parser
                    }
                }

                if (Objects.isNull(offsetTimeValue)) {
                    try {
                        offsetTimeValue = OffsetTime.parse(stringValue);
                    } catch (DateTimeParseException isoParseException) {
                        throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(stringValue), "OffsetTime"), isoParseException);
                    }
                }
            }
        } else if (object instanceof OffsetTime) {
            offsetTimeValue = (OffsetTime) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "OffsetTime"));
        }

        if (Objects.isNull(offsetTimeValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else {
            final String cellString = makeOffsetTimeExportString(offsetTimeValue, columnMeta);
            if (Objects.nonNull(cell)) {
                if (columnMeta.isExportedToString()) {
                    cell.setCellValue(cellString);
                } else {
                    // When there is no pattern/masking, write it as a Numeric (Excel time fraction) cell.
                    // Excel serials have no notion of offset, so only the wall-clock (local) part is stored (the same loss as string export).
                    PxlDateCellSupport.writeNumericTimeCell(cell, offsetTimeValue.toLocalTime(), columnMeta, PxlCodecConstants.localTimeExcelFormat);
                }
            }

            return cellString;
        }
    }

    /**
     * Renders the export string for an {@link OffsetTime}: formats it with the configured export {@link DateTimeFormatter}
     * (or {@link DateTimeFormatter#ISO_OFFSET_TIME} when none is set), then applies string-level export processing via
     * {@link PxlStringCodec#makeExportString}.
     *
     * @param offsetTimeValue the value to render
     * @param columnMeta      resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the value cannot be formatted with the export formatter
     */
    private static String makeOffsetTimeExportString(final OffsetTime offsetTimeValue,
                                                     final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(offsetTimeValue)) {
            return null;
        }

        final DateTimeFormatter exportTimeFormatter = columnMeta.getExportDateTimeFormatterCache();
        try {
            final String stringValue = Objects.nonNull(exportTimeFormatter)
                    ? offsetTimeValue.format(exportTimeFormatter)
                    : offsetTimeValue.format(DateTimeFormatter.ISO_OFFSET_TIME);

            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } catch (DateTimeException dateTimeException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_EXPORT_PARSE_INVALID, String.valueOf(offsetTimeValue), "OffsetTime"), dateTimeException);
        }
    }

}
