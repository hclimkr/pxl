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
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Objects;
import java.util.Optional;

/**
 * Codec for {@link java.time.LocalDateTime} column values — parses cells/strings into {@code LocalDateTime}
 * on import and writes {@code LocalDateTime} into cells on export.
 *
 * <p>Import reads date-formatted NUMERIC cells via POI (other numerics as Excel
 * serials) and strings via the column's cached {@link DateTimeFormatter} (falling back to the built-in
 * read formatters, then ISO-8601). A BOOLEAN cell is rejected as an unsupported cell type. Export writes
 * either a formatted string or a numeric Excel-date cell.
 */
final class PxlLocalDateTimeCodec {

    /**
     * Prevents instantiation.
     */
    private PxlLocalDateTimeCodec() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Parses the given cell into a {@link java.time.LocalDateTime}. Date-formatted NUMERIC cells use POI's
     * local date-time; other numerics are treated as Excel date serials; STRING cells
     * are delegated to the string parser; BOOLEAN cells are rejected as an unsupported cell type; BLANK cells
     * yield {@code null}.
     *
     * @param cell       the cell to read
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@code LocalDateTime}, or {@code null} when blank
     * @throws PxlCellCodecException if the cell type is unsupported or the numeric value is invalid
     */
    static LocalDateTime parseLocalDateTimeValue(final Cell cell,
                                                 final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        LocalDateTime localDateTimeValue = null;

        final CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                try {
                    localDateTimeValue = PxlDateCellSupport.readNumericCellAsLocalDateTime(cell, "LocalDateTime");
                } catch (NumberFormatException numberFormatException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(cell), "LocalDateTime"), numberFormatException);
                }
                break;

            case STRING:
                final String stringCellValue = cell.getStringCellValue();
                localDateTimeValue = parseLocalDateTimeValue(stringCellValue, columnMeta);
                break;

            case BOOLEAN:
                // final boolean booleanCellValue = cell.getBooleanCellValue();
                // localDateTimeValue = DateUtil.getLocalDateTime(BooleanUtils.toInteger(booleanCellValue));
                // break;
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));

            case BLANK:
                // empty
                break;

            default:
                throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CELL_TYPE_UNSUPPORTED, String.valueOf(cellType.toString())));
        }

        return localDateTimeValue;
    }

    /**
     * Parses a string token into a {@link java.time.LocalDateTime}. The column's cached formatter is tried
     * first, then the built-in read formatters, then ISO-8601. The value is trimmed when {@code importTrim}
     * is set; a blank value yields {@code null}.
     *
     * @param s          the raw string token
     * @param columnMeta the resolved import metadata for this column
     * @return the parsed {@code LocalDateTime}, or {@code null} when blank
     * @throws PxlCellCodecException if the value matches no known date-time format
     */
    static LocalDateTime parseLocalDateTimeValue(final String s,
                                                 final PxlImportColumnMeta columnMeta)
            throws PxlCellCodecException {

        final String stringValue = columnMeta.isImportTrim() ? StringUtils.trim(s) : s;
        if (StringUtils.isBlank(stringValue)) {
            return null;
        }

        LocalDateTime localDateTimeValue = null;

        final DateTimeFormatter importDateTimeFormatter = columnMeta.getImportDateTimeFormatterCache();
        if (Objects.nonNull(importDateTimeFormatter)) {
            try {
                localDateTimeValue = LocalDateTime.parse(stringValue, importDateTimeFormatter);
                return localDateTimeValue;
            } catch (DateTimeParseException dateTimeParseException) {
                // go to next parser
            }
        }

        for (final DateTimeFormatter localDateTimeFormatter : PxlCodecConstants.localDateTimeReadFormatters) {
            try {
                localDateTimeValue = LocalDateTime.parse(stringValue, localDateTimeFormatter);
                return localDateTimeValue;
            } catch (DateTimeParseException dateTimeParseException) {
                // go to next parser
            }
        }

        try {
            localDateTimeValue = LocalDateTime.parse(stringValue);
            return localDateTimeValue;
        } catch (DateTimeParseException e) {
            // go to next parser
        }

        throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "LocalDateTime"));
    }

    /**
     * Writes the given value as a {@code LocalDateTime} cell and returns the exported string. A
     * {@code String} source is parsed with the export formatter; a {@link java.time.LocalDateTime} source
     * is used directly. A {@code null} result blanks the cell; otherwise the cell is written as a formatted
     * string when exported to string, or as a numeric Excel-date cell when no pattern/masking applies.
     *
     * @param cell       the target cell, or {@code null} to only compute the string
     * @param object     the source value (a {@code String} or {@code LocalDateTime})
     * @param columnMeta the resolved export metadata for this column
     * @return the exported string, or {@code null} when blank
     * @throws PxlCellCodecException if the source is unsupported or the string is not a valid date-time
     */
    static String buildLocalDateTimeCell(final Cell cell,
                                         final Object object,
                                         final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        DateTimeFormatter exportDateTimeFormatter = columnMeta.getExportDateTimeFormatterCache();
        if (Objects.isNull(exportDateTimeFormatter)) {
            exportDateTimeFormatter = PxlCodecConstants.localDateTimeWriteFormatter;
        }

        LocalDateTime localDateTimeValue;

        if (object instanceof String) {
            final String stringValue = (String) object;

            if (StringUtils.isBlank(stringValue)) {
                localDateTimeValue = null;
            } else {
                try {
                    localDateTimeValue = LocalDateTime.parse(stringValue, exportDateTimeFormatter);
                } catch (DateTimeParseException dateTimeParseException) {
                    throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(stringValue), "LocalDateTime"), dateTimeParseException);
                }
            }
        } else if (object instanceof LocalDateTime) {
            localDateTimeValue = (LocalDateTime) object;
        } else {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_CONVERT_UNSUPPORTED, object.getClass().getSimpleName(), "LocalDateTime"));
        }

        if (Objects.isNull(localDateTimeValue)) {
            Optional.ofNullable(cell).ifPresent(Cell::setBlank);
            return null;
        } else {
            final String cellString = makeLocalDateTimeExportString(localDateTimeValue, columnMeta);
            if (Objects.nonNull(cell)) {
                if (columnMeta.isExportedToString()) {
                    cell.setCellValue(cellString);
                } else {
                    // When there is no pattern/masking, write it as a Numeric (Excel date serial) cell rather than a string.
                    PxlDateCellSupport.writeNumericCell(cell, localDateTimeValue, columnMeta, PxlCodecConstants.localDateTimeExcelFormat);
                }
            }

            return cellString;
        }
    }

    /**
     * Renders the export string for a {@code LocalDateTime}: formats it with the configured export {@link DateTimeFormatter}
     * (or the built-in default write formatter when none is set), then applies string-level export processing via
     * {@link PxlStringCodec#makeExportString}.
     *
     * @param localDateTimeValue the value to render
     * @param columnMeta         resolved export metadata for the column
     * @return the export string representation, or {@code null} when the value is {@code null}
     * @throws PxlCellCodecException if the value cannot be formatted with the export formatter
     */
    private static String makeLocalDateTimeExportString(final LocalDateTime localDateTimeValue,
                                                        final PxlExportColumnMeta columnMeta)
            throws PxlCellCodecException {

        if (Objects.isNull(localDateTimeValue)) {
            return null;
        }

        DateTimeFormatter exportDateTimeFormatter = columnMeta.getExportDateTimeFormatterCache();
        if (Objects.isNull(exportDateTimeFormatter)) {
            exportDateTimeFormatter = PxlCodecConstants.localDateTimeWriteFormatter;
        }

        try {
            final String stringValue = localDateTimeValue.format(exportDateTimeFormatter);

            return PxlStringCodec.makeExportString(stringValue, columnMeta);
        } catch (DateTimeException dateTimeException) {
            throw new PxlCellCodecException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CODEC_PARSE_INVALID, String.valueOf(localDateTimeValue), "LocalDateTime"), dateTimeException);
        }
    }

}
