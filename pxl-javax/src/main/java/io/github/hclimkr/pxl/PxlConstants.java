package io.github.hclimkr.pxl;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;
import io.github.hclimkr.pxl.styler.PxlStyler;
import io.github.hclimkr.pxl.styler.data.PxlDataVerticalCenterTextStyler;
import io.github.hclimkr.pxl.styler.header.PxlHeaderOptionalStyler;
import io.github.hclimkr.pxl.styler.header.PxlHeaderRequiredStyler;
import io.github.hclimkr.pxl.type.PxlExcelEngine;
import io.github.hclimkr.pxl.type.PxlFileFormat;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.lang3.BooleanUtils;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * Central registry of library-wide default values and limits: creator/application metadata, filename
 * extensions and MIME content types, default header-cell styling, and the default import/export/sheet/column
 * options that annotations and {@code option/} overrides fall back to.
 */
public interface PxlConstants {

    /*
     * Component Features
     */


    /*
     * General Constants
     */
    /**
     * Author/creator metadata written into generated workbooks.
     */
    String PXL_CREATOR = "Pxl";
    /**
     * Application metadata written into generated workbooks.
     */
    String PXL_APPLICATION = "Pxl Application";

    /**
     * Filename extension for legacy binary Excel ({@code .xls}).
     */
    String FILENAME_EXTENSION_XLS = "xls";
    /**
     * Filename extension for OOXML Excel ({@code .xlsx}).
     */
    String FILENAME_EXTENSION_XLSX = "xlsx";
    /**
     * Filename extension for CSV ({@code .csv}).
     */
    String FILENAME_EXTENSION_CSV = "csv";

    /**
     * MIME content type for legacy binary Excel ({@code .xls}).
     */
    String CONTENT_TYPE_MICROSOFT_XLS = "application/vnd.ms-excel";
    /**
     * MIME content type for OOXML Excel ({@code .xlsx}).
     */
    String CONTENT_TYPE_MICROSOFT_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    /**
     * Alternate {@code .xlsx} MIME content type emitted by Hancom Office.
     */
    String CONTENT_TYPE_HANCOM_XLSX = "application/haansoftxlsx";
    /**
     * MIME content type for CSV.
     */
    String CONTENT_TYPE_CSV = "text/csv";

    /**
     * Background (foreground fill) color applied to header cells.
     */
    IndexedColors HEADER_COLUMN_FOREGROUND_COLOR = IndexedColors.GREY_25_PERCENT;
    /**
     * Fill pattern applied to header cells.
     */
    FillPatternType HEADER_COLUMN_FILL_PATTERN = FillPatternType.SOLID_FOREGROUND;

    /**
     * Font color for required-column header cells.
     */
    IndexedColors REQUIRED_HEADER_COLUMN_FONT_COLOR = IndexedColors.BLACK;
    /**
     * Background color for required-column header cells.
     */
    IndexedColors REQUIRED_HEADER_COLUMN_FOREGROUND_COLOR = HEADER_COLUMN_FOREGROUND_COLOR;
    /**
     * Fill pattern for required-column header cells.
     */
    FillPatternType REQUIRED_HEADER_COLUMN_FILL_PATTERN = HEADER_COLUMN_FILL_PATTERN;

    /**
     * Font color for optional-column header cells.
     */
    IndexedColors OPTIONAL_HEADER_COLUMN_FONT_COLOR = IndexedColors.GREY_50_PERCENT;
    /**
     * Background color for optional-column header cells.
     */
    IndexedColors OPTIONAL_HEADER_COLUMN_FOREGROUND_COLOR = HEADER_COLUMN_FOREGROUND_COLOR;
    /**
     * Fill pattern for optional-column header cells.
     */
    FillPatternType OPTIONAL_HEADER_COLUMN_FILL_PATTERN = HEADER_COLUMN_FILL_PATTERN;


    /*
     * Workbook Options
     */
    /**
     * Default password expected for encrypted import ({@code ""} means unencrypted).
     */
    String DEFAULT_IMPORT_PASSWORD = "";

    /**
     * Whether bean-validation of imported rows is enabled by default.
     */
    boolean DEFAULT_IMPORT_DATA_VALIDATION = true;

    /**
     * Whether the low-memory streaming reader is used for import by default (XLSX only).
     *
     * @see <a href="https://github.com/pjfanning/excel-streaming-reader">excel-streaming-reader</a>
     */
    boolean DEFAULT_IMPORT_USING_STREAM_READER = false;

    // The number of rows to keep in memory at any given point (PXL default: 100),
    // if DEFAULT_IMPORT_USING_STREAM_READER is enabled.
    /**
     * Default streaming-reader row cache size (rows kept in memory) when streaming import is enabled.
     */
    int DEFAULT_IMPORT_STREAM_READER_ROW_CACHE_SIZE = 100;

    // The number of bytes to read into memory from the input resource (Defaults to 1024),
    // if DEFAULT_IMPORT_USING_STREAM_READER is enabled.
    /**
     * Default streaming-reader input buffer size in bytes when streaming import is enabled.
     */
    int DEFAULT_IMPORT_STREAM_READER_BUFFER_SIZE = 4096;

    /**
     * Default POI engine used to write an Excel workbook on export.
     */
    PxlExcelEngine DEFAULT_EXPORT_EXCEL_ENGINE = PxlExcelEngine.XSSF;

    /**
     * Default physical file format produced by export - the format
     * {@link #DEFAULT_EXPORT_EXCEL_ENGINE} writes, and the fallback of the
     * {@link PxlFileFormat#fromPoiWorkbook(Workbook)} lookup.
     */
    PxlFileFormat DEFAULT_EXPORT_FILE_FORMAT = PxlFileFormat.XLSX;

    /**
     * Default password used to encrypt exported files ({@code ""} means no encryption).
     */
    String DEFAULT_EXPORT_PASSWORD = "";

    /**
     * Whether data-validation dropdowns are written to exported sheets by default.
     */
    boolean DEFAULT_EXPORT_DATA_VALIDATION = true;

    /**
     * SXSSF window size sentinel: keep all rows accessible in memory (no flushing).
     */
    int EXPORT_UNLIMITED_ROW_ACCESS_WINDOW_SIZE = -1;
    /**
     * SXSSF window size sentinel: flush every row immediately (no rows kept accessible).
     */
    int EXPORT_NO_ROW_ACCESS_WINDOW_SIZE = 0;
    /**
     * Default SXSSF row-access window size (rows kept in memory during streaming export).
     */
    int DEFAULT_EXPORT_SXSSF_ROW_ACCESS_WINDOW_SIZE = SXSSFWorkbook.DEFAULT_WINDOW_SIZE;

    /**
     * Maximum number of sheets allowed in an Excel workbook on import.
     */
    int IMPORT_MAX_NUMBER_OF_EXCEL_SHEETS = 100;    // max number of sheets in the Excel to import
    /**
     * Maximum number of sheets allowed in an Excel workbook on export.
     */
    int EXPORT_MAX_NUMBER_OF_EXCEL_SHEETS = 100;    // max number of sheets in the Excel to export

    /**
     * Maximum number of CSV files (sheets) allowed on import.
     */
    int IMPORT_MAX_NUMBER_OF_CSV_SHEETS = 100;      // max number of CSVs to import
    /**
     * Maximum number of rows allowed per CSV on import.
     */
    int IMPORT_MAX_NUMBER_OF_CSV_ROWS = 100_000;    // max number of rows in the CSV to import
    /**
     * Maximum number of columns allowed per CSV on import.
     * <p>
     * CSV itself bounds neither rows nor columns, so this is a defensive ceiling rather than a format rule. It
     * matches {@link PxlFileFormat#XLSX}, which keeps a row class that exports to XLSX readable back from CSV -
     * a lower cap would reject the very files PXL had written.
     */
    int IMPORT_MAX_NUMBER_OF_CSV_COLUMNS = 16_384;  // max number of columns in the CSV to import (= XLSX)
    /**
     * Maximum number of CSV files (sheets) allowed on export.
     */
    int EXPORT_MAX_NUMBER_OF_CSV_SHEETS = 100;      // max number of CSVs to export
    /**
     * Maximum number of rows allowed per CSV on export.
     */
    int EXPORT_MAX_NUMBER_OF_CSV_ROWS = 100_000;    // max number of rows in the CSV to export
    /**
     * Maximum number of columns allowed per CSV on export.
     * <p>
     * Kept equal to {@link #IMPORT_MAX_NUMBER_OF_CSV_COLUMNS} so a CSV PXL writes is one PXL can read back.
     */
    int EXPORT_MAX_NUMBER_OF_CSV_COLUMNS = 16_384;  // max number of columns in the CSV to export (= XLSX)
    /**
     * How many bytes of a CSV export are held in memory before the rest spills to a temporary file.
     * <p>
     * A CSV export renders its whole output before the destination is opened, which is what keeps a failure from
     * leaving a file behind. Holding all of it in memory made the heap the limit: the growth copy of the buffer puts
     * the peak at two to three times the output, so roughly half the heap was the real ceiling. Above this threshold
     * the output continues into a temporary file instead, which trades that ceiling for disk. Below it nothing
     * touches the file system, so the common case - a report of a few thousand rows - behaves exactly as before.
     * <p>
     * The row and column caps above do not bound the output in bytes, so they are no substitute for this.
     */
    int EXPORT_MEMORY_THRESHOLD_OF_CSV = 4 * 1024 * 1024;  // 4 MiB, then spill to a temporary file

    /**
     * Default Commons-CSV parse format used for import (Excel dialect, quoted, empty lines ignored).
     */
    CSVFormat DEFAULT_IMPORT_CSV_FORMAT = CSVFormat
            .EXCEL
            .builder()
            .setSkipHeaderRecord(false)
//            .setIgnoreHeaderCase(true)
            .setIgnoreEmptyLines(true)
            .setQuote('"')
            .setTrim(false)     // !DEFAULT_IMPORT_TRIM
            .build();

    /**
     * Charset used to decode CSV input when no level of the cascade specifies one.
     *
     * @see #UNSPECIFIED_IMPORT_CSV_CHARSET
     * @see <a href="https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html">Java supported encodings</a>
     */
    String DEFAULT_IMPORT_CSV_CHARSET = "UTF-8";

    /**
     * CSV field delimiter used when no level of the cascade specifies one.
     *
     * @see #UNSPECIFIED_IMPORT_CSV_DELIMITER
     */
    char DEFAULT_IMPORT_CSV_DELIMITER = ',';

    /**
     * Marks the CSV charset as not specified at an annotation level on import.
     * <p>
     * A CSV workbook is read as one file per sheet, so the charset resolves through a cascade - sheet option,
     * {@link PxlSheet#importCsvCharset()}, workbook option, {@link PxlWorkbook#importCsvCharset()}, and finally
     * {@link #DEFAULT_IMPORT_CSV_CHARSET}. An option level says "not specified" with {@code null}, which an annotation
     * element cannot hold, so the two annotation levels say it with this value instead. Any blank value counts.
     */
    String UNSPECIFIED_IMPORT_CSV_CHARSET = "";

    /**
     * Marks the CSV field delimiter as not specified at an annotation level on import.
     * <p>
     * Resolves through the same cascade as {@link #UNSPECIFIED_IMPORT_CSV_CHARSET}, ending at
     * {@link #DEFAULT_IMPORT_CSV_DELIMITER}. NUL stands in for "not specified" because it is not a delimiter
     * anyone writes.
     */
    char UNSPECIFIED_IMPORT_CSV_DELIMITER = '\0';

    /**
     * Default Commons-CSV print format used for export (Excel dialect, quoted).
     * <p>
     * A header is deliberately not set on this format: a {@code CSVPrinter} built from a format carrying one
     * writes that header itself on construction, and PXL writes the header from the column metadata, so the
     * file would carry two header lines.
     */
    CSVFormat DEFAULT_EXPORT_CSV_FORMAT = CSVFormat
            .EXCEL
            .builder()
            .setQuote('"')
            .build();

    /**
     * Charset used to encode CSV output when no level of the cascade specifies one.
     *
     * @see #UNSPECIFIED_EXPORT_CSV_CHARSET
     * @see <a href="https://docs.oracle.com/javase/8/docs/technotes/guides/intl/encoding.doc.html">Java supported encodings</a>
     */
    String DEFAULT_EXPORT_CSV_CHARSET = "UTF-8";

    /**
     * CSV field delimiter used when no level of the cascade specifies one.
     *
     * @see #UNSPECIFIED_EXPORT_CSV_DELIMITER
     */
    char DEFAULT_EXPORT_CSV_DELIMITER = ',';

    /**
     * Whether a byte order mark is written ahead of CSV output when nothing specifies otherwise.
     * <p>
     * A BOM is only ever written for UTF-8, UTF-16LE and UTF-16BE. Any other charset leaves it out, including
     * the endian-detecting UTF-16 (whose encoder writes one itself) and the non-Unicode charsets (which cannot
     * encode U+FEFF and would corrupt the first field).
     */
    boolean DEFAULT_EXPORT_CSV_BOM = false;

    /**
     * Marks the CSV charset as not specified at an annotation level on export.
     * <p>
     * A CSV workbook is written as one file per sheet, so the charset resolves through a cascade - sheet option,
     * {@link PxlSheet#exportCsvCharset()}, workbook option, {@link PxlWorkbook#exportCsvCharset()}, and finally
     * {@link #DEFAULT_EXPORT_CSV_CHARSET}. An option level says "not specified" with {@code null}, which an annotation
     * element cannot hold, so the two annotation levels say it with this value instead. Any blank value counts.
     */
    String UNSPECIFIED_EXPORT_CSV_CHARSET = "";

    /**
     * Marks the CSV field delimiter as not specified at an annotation level on export.
     * <p>
     * Resolves through the same cascade as {@link #UNSPECIFIED_EXPORT_CSV_CHARSET}, ending at
     * {@link #DEFAULT_EXPORT_CSV_DELIMITER}. NUL stands in for "not specified" because it is not a delimiter
     * anyone writes.
     */
    char UNSPECIFIED_EXPORT_CSV_DELIMITER = '\0';

    // BASE_NAME: "messages" if the messages_xx_XX.properties files are in src/main/resources
    //            "messages.messages" if the messages_xx_XX.properties files are in src/main/resources/messages
    /**
     * Default i18n resource-bundle base name for import; empty disables i18n (opt-in).
     */
    String DEFAULT_IMPORT_I18N_BASE_NAME = "";  // empty disables i18n (opt-in). To enable, specify a base name (e.g. messages).
    /**
     * Default i18n language tag for import.
     */
    String DEFAULT_IMPORT_I18N_LANGUAGE = "en";
    /**
     * Default i18n country tag for import (empty means none).
     */
    String DEFAULT_IMPORT_I18N_COUNTRY = "";
    /**
     * Default i18n resource-bundle base name for export; empty disables i18n (opt-in).
     */
    String DEFAULT_EXPORT_I18N_BASE_NAME = "";  // empty disables i18n (opt-in). To enable, specify a base name (e.g. messages).
    /**
     * Default i18n language tag for export.
     */
    String DEFAULT_EXPORT_I18N_LANGUAGE = "en";
    /**
     * Default i18n country tag for export (empty means none).
     */
    String DEFAULT_EXPORT_I18N_COUNTRY = "";


    /*
     * Sheet Common Options
     */
    /**
     * Wildcard sheet field name that matches any sheet.
     */
    String SHEET_FIELD_NAME_WILD_CARD = "*";
    /**
     * Maximum length of an Excel sheet name (Excel limit).
     */
    int MAX_SHEET_NAME_LENGTH = 31;


    /*
     * Sheet Import Options
     */
    /**
     * Whether a subclass sheet definition overrides its superclass sheet by default on import.
     */
    boolean DEFAULT_IMPORT_OVERRIDE_SUPER_CLASS_SHEET = false;
    /**
     * Whether hidden rows are excluded from import by default.
     */
    boolean DEFAULT_IMPORT_EXCLUDE_HIDDEN_ROWS = false;
    /**
     * Whether hidden columns are excluded from import by default.
     */
    boolean DEFAULT_IMPORT_EXCLUDE_HIDDEN_COLUMNS = false;
    /**
     * Whether every cell of a merged region is read (vs. only the top-left cell) by default.
     */
    boolean DEFAULT_IMPORT_EACH_CELL_OF_MERGED_REGION = false;
    /**
     * Default header row index for import ({@code 0} means auto/unset).
     */
    int DEFAULT_IMPORT_HEADER_ROW_INDEX = 0;
    /**
     * Default first data row index for import ({@code 0} means auto/unset).
     */
    int DEFAULT_IMPORT_FIRST_DATA_ROW_INDEX = 0;
    /**
     * Default last data row index for import ({@code 0} means auto/unset).
     */
    int DEFAULT_IMPORT_LAST_DATA_ROW_INDEX = 0;
    /**
     * Default first data column index for import ({@code 0} means auto/unset).
     */
    int DEFAULT_IMPORT_FIRST_DATA_COLUMN_INDEX = 0;
    /**
     * Default last data column index for import ({@code 0} means auto/unset).
     */
    int DEFAULT_IMPORT_LAST_DATA_COLUMN_INDEX = 0;


    /*
     * Sheet Export Options
     */
    /**
     * Whether a subclass sheet definition overrides its superclass sheet by default on export.
     */
    boolean DEFAULT_EXPORT_OVERRIDE_SUPER_CLASS_SHEET = false;
    /**
     * Default data row height in points ({@code -1} means unset/auto).
     */
    float DEFAULT_EXPORT_ROW_HEIGHT_IN_POINTS = -1.F;
    /**
     * Default header row index for export ({@code 0} means auto/unset).
     */
    int DEFAULT_EXPORT_HEADER_ROW_INDEX = 0;
    /**
     * Default first data row index for export ({@code 0} means auto/unset).
     */
    int DEFAULT_EXPORT_FIRST_DATA_ROW_INDEX = 0;
    /**
     * Default last data row index for export ({@code 0} means auto/unset).
     */
    int DEFAULT_EXPORT_LAST_DATA_ROW_INDEX = 0;
    /**
     * Default first data column index for export ({@code 0} means auto/unset).
     */
    int DEFAULT_EXPORT_FIRST_DATA_COLUMN_INDEX = 0;
    /**
     * Default last data column index for export ({@code 0} means auto/unset).
     */
    int DEFAULT_EXPORT_LAST_DATA_COLUMN_INDEX = 0;

    /**
     * Whether a cell is written when the source value is {@code null} by default.
     */
    boolean DEFAULT_EXPORT_IF_NULL = false;
    /**
     * Whether a cell is written when the source value is empty by default.
     */
    boolean DEFAULT_EXPORT_IF_EMPTY = true;

    /**
     * Whether an auto-filter is applied to the header row by default.
     */
    boolean DEFAULT_EXPORT_COLUMN_FILTER = false;

    /**
     * Sentinel styler type meaning "no styler specified" (falls back to the next level in the cascade).
     */
    Class<? extends PxlStyler> VOID_CELL_STYLER = PxlStyler.class;

    /**
     * Default workbook-level styler for required-column header cells.
     */
    Class<? extends PxlStyler> DEFAULT_EXPORT_WORKBOOK_REQUIRED_HEADER_CELL_STYLER = PxlHeaderRequiredStyler.class;
    /**
     * Default workbook-level styler for optional-column header cells.
     */
    Class<? extends PxlStyler> DEFAULT_EXPORT_WORKBOOK_OPTIONAL_HEADER_CELL_STYLER = PxlHeaderOptionalStyler.class;
    /**
     * Default workbook-level styler for data cells.
     */
    Class<? extends PxlStyler> DEFAULT_EXPORT_WORKBOOK_DATA_CELL_STYLER = PxlDataVerticalCenterTextStyler.class;

    /**
     * Sheet-level default styler for required-column header cells; unspecified by default.
     *
     * @deprecated retained for backward compatibility; configure stylers via the
     * {@code @PxlWorkbook}/{@code @PxlSheet}/{@code @PxlColumn} styler attributes instead.
     */
    @Deprecated
    Class<? extends PxlStyler> DEFAULT_EXPORT_SHEET_REQUIRED_HEADER_CELL_STYLER = VOID_CELL_STYLER;
    /**
     * Sheet-level default styler for optional-column header cells; unspecified by default.
     *
     * @deprecated retained for backward compatibility; configure stylers via the
     * {@code @PxlWorkbook}/{@code @PxlSheet}/{@code @PxlColumn} styler attributes instead.
     */
    @Deprecated
    Class<? extends PxlStyler> DEFAULT_EXPORT_SHEET_OPTIONAL_HEADER_CELL_STYLER = VOID_CELL_STYLER;
    /**
     * Sheet-level default styler for data cells; unspecified by default.
     *
     * @deprecated retained for backward compatibility; configure stylers via the
     * {@code @PxlWorkbook}/{@code @PxlSheet}/{@code @PxlColumn} styler attributes instead.
     */
    @Deprecated
    Class<? extends PxlStyler> DEFAULT_EXPORT_SHEET_DATA_CELL_STYLER = VOID_CELL_STYLER;

    /**
     * Column-level default styler for required-column header cells; unspecified by default.
     *
     * @deprecated retained for backward compatibility; configure stylers via the
     * {@code @PxlWorkbook}/{@code @PxlSheet}/{@code @PxlColumn} styler attributes instead.
     */
    @Deprecated
    Class<? extends PxlStyler> DEFAULT_EXPORT_COLUMN_REQUIRED_HEADER_CELL_STYLER = VOID_CELL_STYLER;
    /**
     * Column-level default styler for optional-column header cells; unspecified by default.
     *
     * @deprecated retained for backward compatibility; configure stylers via the
     * {@code @PxlWorkbook}/{@code @PxlSheet}/{@code @PxlColumn} styler attributes instead.
     */
    @Deprecated
    Class<? extends PxlStyler> DEFAULT_EXPORT_COLUMN_OPTIONAL_HEADER_CELL_STYLER = VOID_CELL_STYLER;
    /**
     * Column-level default styler for data cells; unspecified by default.
     *
     * @deprecated retained for backward compatibility; configure stylers via the
     * {@code @PxlWorkbook}/{@code @PxlSheet}/{@code @PxlColumn} styler attributes instead.
     */
    @Deprecated
    Class<? extends PxlStyler> DEFAULT_EXPORT_COLUMN_DATA_CELL_STYLER = VOID_CELL_STYLER;


    /*
     * Column Common Options
     */
    /**
     * Default separator used to join/split collection elements within a single cell.
     */
    String DEFAULT_COLLECTION_SEPARATOR = ";";


    /*
     * Column Import Options
     */
    /**
     * Whether imported cell strings are trimmed by default.
     */
    boolean DEFAULT_IMPORT_TRIM = true;
    /**
     * Whether column values must be unique across rows by default on import.
     */
    boolean DEFAULT_IMPORT_UNIQUE = false;

    /**
     * Default string parsed as boolean {@code true} on import.
     */
    String DEFAULT_IMPORT_TRUE_STRING = BooleanUtils.TRUE;
    /**
     * Default string parsed as boolean {@code false} on import.
     */
    String DEFAULT_IMPORT_FALSE_STRING = BooleanUtils.FALSE;

    /**
     * Default per-column collection separator for import; empty falls back to {@link #DEFAULT_COLLECTION_SEPARATOR}.
     */
    String DEFAULT_IMPORT_COLLECTION_SEPARATOR = "";

    /**
     * Whether a subclass column definition overrides its superclass column by default on import.
     */
    boolean DEFAULT_IMPORT_OVERRIDE_SUPER_CLASS_COLUMN = false;


    /*
     * Column Export Options
     */
    /**
     * Whether exported cell strings are trimmed by default.
     */
    boolean DEFAULT_EXPORT_TRIM = false;

    /**
     * Column width sentinel: auto-size the column from its content.
     */
    int EXPORT_AUTO_COLUMN_WIDTH = 0;   // too expensive
    /**
     * Lower bound applied when auto-sizing a column width.
     */
    int EXPORT_AUTO_COLUMN_MIN_WIDTH = 2_000;
    /**
     * Upper bound applied when auto-sizing a column width.
     */
    int EXPORT_AUTO_COLUMN_MAX_WIDTH = 15_000;

    /**
     * Column width sentinel: leave the column width unset.
     */
    int EXPORT_NONE_COLUMN_WIDTH = -1;  // don't set column width
    /**
     * Default export column width (auto-size).
     */
    int DEFAULT_EXPORT_COLUMN_WIDTH = EXPORT_AUTO_COLUMN_WIDTH;

    /**
     * Default per-column collection separator for export; empty falls back to {@link #DEFAULT_COLLECTION_SEPARATOR}.
     */
    String DEFAULT_EXPORT_COLLECTION_SEPARATOR = "";

    /**
     * Whether a subclass column definition overrides its superclass column by default on export.
     */
    boolean DEFAULT_EXPORT_OVERRIDE_SUPER_CLASS_COLUMN = false;

    /**
     * Default drop-down list style used when exporting enum columns as data-validation.
     */
    PxlColumn.ExportEnumDropDownListStyle DEFAULT_EXPORT_ENUM_DROP_DOWN_LIST_STYLE = PxlColumn.ExportEnumDropDownListStyle.SET;

    /**
     * Default cell text written for a {@code null} value on export.
     */
    String DEFAULT_EXPORT_NULL_STRING = "";
    /**
     * Default string written for boolean {@code true} on export.
     */
    String DEFAULT_EXPORT_TRUE_STRING = BooleanUtils.TRUE;
    /**
     * Default string written for boolean {@code false} on export.
     */
    String DEFAULT_EXPORT_FALSE_STRING = BooleanUtils.FALSE;

    /**
     * On-screen picture cell width in pixels for exported images.
     */
    int EXPORT_PICTURE_SCREEN_WIDTH_IN_PIXELS = 100;
    /**
     * On-screen picture cell height in pixels for exported images.
     */
    int EXPORT_PICTURE_SCREEN_HEIGHT_IN_PIXELS = 100;
    /**
     * Padding in pixels around exported pictures within a cell.
     */
    int EXPORT_PICTURE_SCREEN_PADDING_IN_PIXELS = 10;
    /**
     * Number of pictures laid out horizontally per cell when exporting a collection of images.
     */
    int EXPORT_HORIZONTAL_NUMBER_OF_PICTURE = 6;

    /**
     * Image-scaling backend used when downscaling pictures for export.
     */
    enum PictureScaler {

        /**
         * No scaling; images are embedded at their original size.
         */
        NO_SCALER,
        /**
         * Scale using the Thumbnailator library.
         */
        THUMBNAILATOR,
        /**
         * Scale using the imgscalr library.
         */
        IMGSCALR,   // seems slightly faster than THUMBNAILATOR.

    }

    /**
     * Default picture-scaling backend for export.
     */
    PictureScaler EXPORT_PICTURE_SCALER = PictureScaler.IMGSCALR;
    /**
     * Target width in pixels to which exported pictures are scaled.
     */
    int EXPORT_PICTURE_SCALE_WIDTH_IN_PIXELS = 500;
    /**
     * Target height in pixels to which exported pictures are scaled.
     */
    int EXPORT_PICTURE_SCALE_HEIGHT_IN_PIXELS = 500;
    /**
     * Whether string values are rendered as pictures on export by default.
     */
    boolean DEFAULT_EXPORT_STRING_AS_PICTURE = false;

    /**
     * Whether string values are written as formulas on export by default.
     */
    boolean DEFAULT_EXPORT_STRING_AS_FORMULA = false;

    /**
     * Default character used to mask values on export.
     */
    String DEFAULT_EXPORT_MASKING_CHAR = "*";

}
