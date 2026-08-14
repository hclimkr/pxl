package io.github.hclimkr.pxl.internal.i18n;

/**
 * Message keys for the PXL library's own diagnostic bundle ({@code pxl-messages}).
 * <p>
 * These keys index {@code pxl-messages*.properties} (shipped inside the pxl artifact) and are resolved
 * through {@link PxlI18nDiagnostic}. They are distinct from the consumer-facing sheet/column name translation,
 * whose bundle and locale come from {@code @PxlWorkbook}; exception/diagnostic text is a library-owned,
 * process-wide concern (see {@link PxlI18nDiagnostic} for the locale strategy).
 */
public final class PxlI18nDiagnosticKeys {

    /**
     * Prevents instantiation.
     */
    private PxlI18nDiagnosticKeys() {

        throw new AssertionError("no instances of this class");
    }

    // Location tag builder (PxlException#buildTagMessage) - each takes a single positional argument,
    // except TAG_JOIN which takes the joined tag and the message.

    /**
     * Location tag prefix for a sheet; {@code {0}}=sheet name.
     */
    public static final String TAG_SHEET = "tag.sheet";
    /**
     * Location tag prefix for a row; {@code {0}}=row number.
     */
    public static final String TAG_ROW = "tag.row";
    /**
     * Location tag prefix for a column identified by name; {@code {0}}=column name.
     */
    public static final String TAG_COLUMN_NAME = "tag.column.name";
    /**
     * Location tag prefix for a column identified by index; {@code {0}}=column index.
     */
    public static final String TAG_COLUMN_INDEX = "tag.column.index";
    /**
     * Joins the location tag and the message; {@code {0}}=tag, {@code {1}}=message.
     */
    public static final String TAG_JOIN = "tag.join";

    // Argument preconditions (PxlAssertSupport). The *_NAMED variants take the parameter name as {0}.

    /**
     * Generic non-null precondition failure (no parameter name).
     */
    public static final String ASSERT_NOT_NULL = "assert.notNull";
    /**
     * Named non-null precondition failure; {@code {0}}=parameter name.
     */
    public static final String ASSERT_NOT_NULL_NAMED = "assert.notNull.named";
    /**
     * Generic non-empty precondition failure (no parameter name).
     */
    public static final String ASSERT_NOT_EMPTY = "assert.notEmpty";
    /**
     * Named non-empty precondition failure; {@code {0}}=parameter name.
     */
    public static final String ASSERT_NOT_EMPTY_NAMED = "assert.notEmpty.named";
    /**
     * Generic non-blank precondition failure (no parameter name).
     */
    public static final String ASSERT_NOT_BLANK = "assert.notBlank";
    /**
     * Named non-blank precondition failure; {@code {0}}=parameter name.
     */
    public static final String ASSERT_NOT_BLANK_NAMED = "assert.notBlank.named";
    /**
     * Boolean condition precondition failure.
     */
    public static final String ASSERT_IS_TRUE = "assert.isTrue";

    // Type codecs (internal/codec). Most are parameterized shared templates reused across ~30 codecs.

    /**
     * Unsupported cell type encountered during codec dispatch; {@code {0}}=cellType.
     */
    public static final String CODEC_IMPORT_CELL_TYPE_UNSUPPORTED = "codec.import.cellType.unsupported";
    /**
     * Value read from a cell cannot be parsed into the target type; {@code {0}}=value, {@code {1}}=typeName.
     */
    public static final String CODEC_IMPORT_PARSE_INVALID = "codec.import.parse.invalid";
    /**
     * String export value cannot be parsed into the target type before being written; {@code {0}}=value,
     * {@code {1}}=typeName.
     */
    public static final String CODEC_EXPORT_PARSE_INVALID = "codec.export.parse.invalid";
    /**
     * Unsupported conversion between two types; {@code {0}}=sourceType, {@code {1}}=targetType.
     */
    public static final String CODEC_EXPORT_CONVERT_UNSUPPORTED = "codec.export.convert.unsupported";
    /**
     * Failed to apply the configured format pattern; {@code {0}}=value.
     */
    public static final String CODEC_EXPORT_PATTERN_APPLY_FAILED = "codec.export.pattern.applyFailed";
    /**
     * Unsupported column type; {@code {0}}=type.
     */
    public static final String CODEC_COLUMN_TYPE_UNSUPPORTED = "codec.columnType.unsupported";
    /**
     * Unsupported collection element type; {@code {0}}=type.
     */
    public static final String CODEC_ELEMENT_TYPE_UNSUPPORTED = "codec.elementType.unsupported";
    /**
     * Formula cell cannot be evaluated by the streaming reader.
     */
    public static final String CODEC_IMPORT_STREAMING_FORMULA_UNSUPPORTED = "codec.import.streaming.formulaUnsupported";
    /**
     * Enum constant not found for a value read from a cell; {@code {0}}=value, {@code {1}}=enumType.
     */
    public static final String CODEC_IMPORT_ENUM_PARSE_FAILED = "codec.import.enum.parseFailed";
    /**
     * Enum constant not found for a string export value; {@code {0}}=value, {@code {1}}=enumType.
     */
    public static final String CODEC_EXPORT_ENUM_PARSE_FAILED = "codec.export.enum.parseFailed";
    /**
     * Error while parsing an enum value read from a cell; {@code {0}}=value, {@code {1}}=enumType.
     */
    public static final String CODEC_IMPORT_ENUM_PARSE_ERROR = "codec.import.enum.parseError";
    /**
     * Error while parsing a string export value into an enum; {@code {0}}=value, {@code {1}}=enumType.
     */
    public static final String CODEC_EXPORT_ENUM_PARSE_ERROR = "codec.export.enum.parseError";
    /**
     * Error while formatting an enum value; {@code {0}}=enumType.
     */
    public static final String CODEC_EXPORT_ENUM_FORMAT_ERROR = "codec.export.enum.formatError";
    /**
     * Custom object parse failed for a value read from a cell; {@code {0}}=value, {@code {1}}=type.
     */
    public static final String CODEC_IMPORT_OBJECT_PARSE_FAILED = "codec.import.object.parseFailed";
    /**
     * Custom object parse failed for a string export value; {@code {0}}=value, {@code {1}}=type.
     */
    public static final String CODEC_EXPORT_OBJECT_PARSE_FAILED = "codec.export.object.parseFailed";
    /**
     * Error while parsing a custom object read from a cell; {@code {0}}=value, {@code {1}}=type.
     */
    public static final String CODEC_IMPORT_OBJECT_PARSE_ERROR = "codec.import.object.parseError";
    /**
     * Error while parsing a string export value into a custom object; {@code {0}}=value, {@code {1}}=type.
     */
    public static final String CODEC_EXPORT_OBJECT_PARSE_ERROR = "codec.export.object.parseError";
    /**
     * Error while formatting a custom object; {@code {0}}=type.
     */
    public static final String CODEC_EXPORT_OBJECT_FORMAT_ERROR = "codec.export.object.formatError";
    /**
     * Numeric value exceeds the type's representable precision; {@code {0}}=value, {@code {1}}=type.
     */
    public static final String CODEC_EXPORT_VALUE_TOO_LARGE = "codec.export.value.tooLarge";

    // Binders / importers / exporter (internal/core).

    /**
     * CSV file count does not match the expected name count.
     */
    public static final String CORE_IMPORT_CSV_FILE_NAME_COUNT_MISMATCH = "core.import.csv.fileNameCountMismatch";
    /**
     * CSV name matches more than one candidate sheet; {@code {0}}=csvName.
     */
    public static final String CORE_IMPORT_CSV_MULTIPLE_SHEET_MATCH = "core.import.csv.multipleSheetMatch";
    /**
     * {@code importCsvCharset} does not name a charset this JVM supports; {@code {0}}=sheetName, {@code {1}}=charsetName.
     */
    public static final String CORE_IMPORT_CSV_CHARSET_INVALID = "core.import.csv.charsetInvalid";
    /**
     * {@code importCsvDelimiter} cannot be used as a CSV delimiter (a line break, or the quote character); {@code {0}}=sheetName.
     */
    public static final String CORE_IMPORT_CSV_DELIMITER_INVALID = "core.import.csv.delimiterInvalid";
    /**
     * Sheet count exceeds the configured maximum; {@code {0}}=max.
     */
    public static final String CORE_SHEET_COUNT_EXCEEDED = "core.sheet.countExceeded";
    /**
     * Duplicate sheet detected; {@code {0}}=sheetName.
     */
    public static final String CORE_IMPORT_SHEET_DUPLICATE = "core.import.sheet.duplicate";
    /**
     * Requested sheet not found; {@code {0}}=sheetName.
     */
    public static final String CORE_IMPORT_SHEET_NOT_FOUND = "core.import.sheet.notFound";
    /**
     * Sheet row count exceeds the configured maximum; {@code {0}}=sheetName, {@code {1}}=max.
     */
    public static final String CORE_IMPORT_SHEET_ROW_COUNT_EXCEEDED = "core.import.sheet.rowCountExceeded";
    /**
     * Sheet column count exceeds the configured maximum; {@code {0}}=sheetName, {@code {1}}=max.
     */
    public static final String CORE_IMPORT_SHEET_COLUMN_COUNT_EXCEEDED = "core.import.sheet.columnCountExceeded";
    /**
     * No header column found in the header row; {@code {0}}=sheetName, {@code {1}}=rowNumber.
     */
    public static final String CORE_IMPORT_SHEET_NO_HEADER_COLUMN = "core.import.sheet.noHeaderColumn";
    /**
     * Header row is missing; {@code {0}}=sheetName, {@code {1}}=rowNumber.
     */
    public static final String CORE_IMPORT_SHEET_NO_HEADER_ROW = "core.import.sheet.noHeaderRow";
    /**
     * Duplicate columns detected; {@code {0}}=sheetName, {@code {1}}=columnNames.
     */
    public static final String CORE_IMPORT_COLUMN_DUPLICATE = "core.import.column.duplicate";
    /**
     * Expected columns not found; {@code {0}}=sheetName, {@code {1}}=columnNames.
     */
    public static final String CORE_IMPORT_COLUMN_NOT_FOUND = "core.import.column.notFound";
    /**
     * Duplicate values in a unique column; {@code {0}}=duplicate values.
     */
    public static final String CORE_IMPORT_COLUMN_VALUE_DUPLICATE = "core.import.column.valueDuplicate";
    /**
     * Unsupported row-index field type; {@code {0}}=fieldName, {@code {1}}=type.
     */
    public static final String CORE_IMPORT_ROW_INDEX_TYPE_UNSUPPORTED = "core.import.rowIndex.typeUnsupported";
    /**
     * Merged regions are not supported while streaming.
     */
    public static final String CORE_IMPORT_STREAMING_MERGED_UNSUPPORTED = "core.import.streaming.mergedUnsupported";
    /**
     * No data supplied to export.
     */
    public static final String CORE_EXPORT_NO_DATA = "core.export.noData";
    /**
     * Sheet-name count does not match the data-object count on export.
     */
    public static final String CORE_EXPORT_SHEET_NAME_OBJECT_COUNT_MISMATCH = "core.export.sheetNameObjectCountMismatch";
    /**
     * Sheet-name count does not match the row-class count on export.
     */
    public static final String CORE_EXPORT_SHEET_NAME_ROW_CLASS_COUNT_MISMATCH = "core.export.sheetNameRowClassCountMismatch";
    /**
     * Duplicate sheet name supplied on export; {@code {0}}=names.
     */
    public static final String CORE_EXPORT_DUPLICATE_SHEET_NAME = "core.export.duplicateSheetName";
    /**
     * Sheet data is null on export; {@code {0}}=sheetName.
     */
    public static final String CORE_EXPORT_SHEET_DATA_NULL = "core.export.sheetDataNull";
    /**
     * Sheet row class is null on export; {@code {0}}=sheetName.
     */
    public static final String CORE_EXPORT_SHEET_ROW_CLASS_NULL = "core.export.sheetRowClassNull";
    /**
     * Export row count exceeds the configured maximum; {@code {0}}=sheetName, {@code {1}}=max.
     */
    public static final String CORE_EXPORT_SHEET_ROW_COUNT_EXCEEDED = "core.export.sheetRowCountExceeded";
    /**
     * A row object to export is {@code null}; {@code {0}}=sheetName, {@code {1}}=one-based position in the row collection.
     */
    public static final String CORE_EXPORT_ROW_NULL = "core.export.rowNull";

    // Resolved metadata (internal/meta).

    /**
     * Sheet index must be non-negative; {@code {0}}=sheetNames, {@code {1}}=indexName.
     */
    public static final String META_SHEET_INDEX_NON_NEGATIVE = "meta.sheet.indexNonNegative";
    /**
     * First data row must be greater than the header row; {@code {0}}=sheetNames, {@code {1}}=firstDataName, {@code {2}}=headerName.
     */
    public static final String META_SHEET_FIRST_DATA_ROW_GT_HEADER = "meta.sheet.firstDataRowGtHeader";
    /**
     * Last data row must be greater than or equal to the first data row; {@code {0}}=sheetNames, {@code {1}}=lastName, {@code {2}}=firstName.
     */
    public static final String META_SHEET_LAST_DATA_GE_FIRST_DATA = "meta.sheet.lastDataGeFirstData";
    /**
     * Grouping field not found on the row class; {@code {0}}=fieldName, {@code {1}}=rowClass.
     */
    public static final String META_EXPORT_SHEET_GROUPING_FIELD_NOT_FOUND = "meta.export.sheet.groupingFieldNotFound";
    /**
     * Duplicate sheet name; {@code {0}}=names.
     */
    public static final String META_EXPORT_DUPLICATE_SHEET_NAME = "meta.export.duplicateSheetName";
    /**
     * Declared type is not a collection type; {@code {0}}=type.
     */
    public static final String META_NOT_COLLECTION_TYPE = "meta.notCollectionType";
    /**
     * Sheet name is invalid; {@code {0}}=sheetName.
     */
    public static final String META_SHEET_NAME_INVALID = "meta.sheetNameInvalid";
    /**
     * Unsupported column type; {@code {0}}=type.
     */
    public static final String META_COLUMN_TYPE_UNSUPPORTED = "meta.columnTypeUnsupported";
    /**
     * Column count exceeds the configured maximum; {@code {0}}=sheetName, {@code {1}}=max.
     */
    public static final String META_EXPORT_COLUMN_COUNT_EXCEEDED = "meta.export.column.countExceeded";
    /**
     * Column range was truncated to fit the sheet limit; {@code {0}}=sheetName.
     */
    public static final String META_EXPORT_COLUMN_RANGE_TRUNCATED = "meta.export.column.rangeTruncated";
    /**
     * Duplicate column name; {@code {0}}=sheetName, {@code {1}}=names.
     */
    public static final String META_EXPORT_DUPLICATE_COLUMN_NAME = "meta.export.duplicateColumnName";
    /**
     * Invalid export number pattern; {@code {0}}=pattern.
     */
    public static final String META_EXPORT_NUMBER_PATTERN_INVALID = "meta.export.numberPatternInvalid";
    /**
     * Invalid export date/time pattern; {@code {0}}=pattern.
     */
    public static final String META_EXPORT_DATE_TIME_PATTERN_INVALID = "meta.export.dateTimePatternInvalid";
    /**
     * Invalid import number pattern; {@code {0}}=pattern.
     */
    public static final String META_IMPORT_NUMBER_PATTERN_INVALID = "meta.import.numberPatternInvalid";
    /**
     * Invalid import date/time pattern; {@code {0}}=pattern.
     */
    public static final String META_IMPORT_DATE_TIME_PATTERN_INVALID = "meta.import.dateTimePatternInvalid";
    /**
     * Invalid masking regex; {@code {0}}=regex.
     */
    public static final String META_EXPORT_MASKING_INVALID = "meta.export.maskingInvalid";
    /**
     * Export converter must return a String; {@code {0}}=type.
     */
    public static final String META_EXPORT_CONVERTER_RETURN_STRING = "meta.export.converterReturnString";
    /**
     * Static export converter must take exactly one parameter; {@code {0}}=type.
     */
    public static final String META_EXPORT_CONVERTER_STATIC_ONE_PARAM = "meta.export.converterStaticOneParam";
    /**
     * Instance export converter must take no parameter; {@code {0}}=type.
     */
    public static final String META_EXPORT_CONVERTER_INSTANCE_NO_PARAM = "meta.export.converterInstanceNoParam";
    /**
     * Import converter return type does not match the field type; {@code {0}}=type.
     */
    public static final String META_IMPORT_CONVERTER_RETURN_TYPE = "meta.import.converterReturnType";
    /**
     * Import converter must be static; {@code {0}}=type.
     */
    public static final String META_IMPORT_CONVERTER_STATIC = "meta.import.converterStatic";
    /**
     * Import converter must take exactly one String parameter; {@code {0}}=type.
     */
    public static final String META_IMPORT_CONVERTER_ONE_STRING_PARAM = "meta.import.converterOneStringParam";

    // Internal plumbing helpers (internal/support).

    /**
     * Number is out of the type's valid range; {@code {0}}=value, {@code {1}}=type, {@code {2}}=min, {@code {3}}=max.
     */
    public static final String SUPPORT_NUMBER_OUT_OF_RANGE = "support.number.outOfRange";
    /**
     * NaN read where a finite number is required; {@code {0}}=type, {@code {1}}=value.
     */
    public static final String SUPPORT_NUMBER_NAN_READ = "support.number.nanRead";
    /**
     * NaN cannot be written to a numeric cell; {@code {0}}=value.
     */
    public static final String SUPPORT_NUMBER_NAN_WRITE = "support.number.nanWrite";
    /**
     * Unsupported type; {@code {0}}=type.
     */
    public static final String SUPPORT_TYPE_UNSUPPORTED = "support.type.unsupported";
    /**
     * Reflective instantiation failed; {@code {0}}=type.
     */
    public static final String SUPPORT_REFLECT_INSTANTIATE_ERROR = "support.reflect.instantiateError";
    /**
     * Type has no no-argument constructor; {@code {0}}=type.
     */
    public static final String SUPPORT_REFLECT_NO_NO_ARG_CONSTRUCTOR = "support.reflect.noNoArgConstructor";
    /**
     * Annotation member is missing; {@code {0}}=annotationType, {@code {1}}=member.
     */
    public static final String SUPPORT_REFLECT_MEMBER_MISSING = "support.reflect.memberMissing";
    /**
     * Annotation member has an incompatible return type; {@code {0}}=member, {@code {1}}=returnType.
     */
    public static final String SUPPORT_REFLECT_MEMBER_INCOMPATIBLE = "support.reflect.memberIncompatible";
    /**
     * Field uses a raw (non-parameterized) collection type; {@code {0}}=fieldName.
     */
    public static final String SUPPORT_REFLECT_RAW_TYPE = "support.reflect.rawType";
    /**
     * Unsupported generic element type; {@code {0}}=fieldName, {@code {1}}=typeName.
     */
    public static final String SUPPORT_REFLECT_ELEMENT_TYPE_UNSUPPORTED = "support.reflect.elementTypeUnsupported";
    /**
     * Workbook-name field must be of type String; {@code {0}}=fieldName, {@code {1}}=type.
     */
    public static final String SUPPORT_WORKBOOK_NAME_STRING_ONLY = "support.workbook.nameStringOnly";
    /**
     * Unsupported workbook format.
     */
    public static final String SUPPORT_WORKBOOK_UNSUPPORTED_FORMAT = "support.workbook.unsupportedFormat";
    /**
     * Workbook decryption failed.
     */
    public static final String SUPPORT_WORKBOOK_DECRYPT_FAILED = "support.workbook.decryptFailed";
    /**
     * Workbook file not found; {@code {0}}=detail.
     */
    public static final String SUPPORT_WORKBOOK_FILE_NOT_FOUND = "support.workbook.fileNotFound";
    /**
     * Workbook file cannot be read; {@code {0}}=detail.
     */
    public static final String SUPPORT_WORKBOOK_FILE_UNREADABLE = "support.workbook.fileUnreadable";
    /**
     * Streaming workbook cannot be read.
     */
    public static final String SUPPORT_WORKBOOK_STREAMING_UNREADABLE = "support.workbook.streamingUnreadable";
    /**
     * Duration component is out of range; {@code {0}}=value.
     */
    public static final String SUPPORT_PERIOD_DURATION_DURATION_RANGE = "support.periodDuration.durationRange";
    /**
     * Period integer component is out of range; {@code {0}}=value.
     */
    public static final String SUPPORT_PERIOD_DURATION_INT_RANGE = "support.periodDuration.intRange";
    /**
     * Period/duration value does not match the expected pattern; {@code {0}}=value, {@code {1}}=pattern.
     */
    public static final String SUPPORT_PERIOD_DURATION_PATTERN_MISMATCH = "support.periodDuration.patternMismatch";
    /**
     * Type reference is missing its type parameter.
     */
    public static final String SUPPORT_TYPE_REFERENCE_MISSING = "support.typeReference.missingTypeParam";

    // Public POI helpers (util).

    /**
     * Image cannot be read from the source; {@code {0}}=url.
     */
    public static final String UTIL_IMAGE_UNREADABLE = "util.image.unreadable";
    /**
     * Invalid cell reference; {@code {0}}=cellRef.
     */
    public static final String UTIL_CELL_REF_INVALID = "util.cellRef.invalid";
    /**
     * Workbook decryption failed.
     */
    public static final String UTIL_WORKBOOK_DECRYPT_FAILED = "util.workbook.decryptFailed";
    /**
     * Unsupported workbook format.
     */
    public static final String UTIL_WORKBOOK_UNSUPPORTED_FORMAT = "util.workbook.unsupportedFormat";
    /**
     * Workbook file not found.
     */
    public static final String UTIL_WORKBOOK_FILE_NOT_FOUND = "util.workbook.fileNotFound";
    /**
     * Workbook file cannot be read.
     */
    public static final String UTIL_WORKBOOK_FILE_UNREADABLE = "util.workbook.fileUnreadable";
    /**
     * Workbook file cannot be opened.
     */
    public static final String UTIL_WORKBOOK_FILE_UNOPENABLE = "util.workbook.fileUnopenable";
    /**
     * Encryption is not supported for the workbook type; {@code {0}}=workbookType.
     */
    public static final String UTIL_WORKBOOK_ENCRYPT_UNSUPPORTED = "util.workbook.encryptUnsupported";

    // Fluent builders (builder).

    /**
     * Workbook and sheet forms are mutually exclusive; {@code {0}}=workbook form.
     */
    public static final String BUILDER_EXPORT_WORKBOOK_SHEET_EXCLUSIVE = "builder.export.workbookSheetExclusive";
    /**
     * Either the workbook or the sheet form is required; {@code {0}}=workbook form.
     */
    public static final String BUILDER_EXPORT_WORKBOOK_SHEET_REQUIRED = "builder.export.workbookSheetRequired";
    /**
     * A CSV terminal writes a single sheet only.
     */
    public static final String BUILDER_EXPORT_CSV_SINGLE_SHEET_ONLY = "builder.export.csv.singleSheetOnly";
    /**
     * CSV output cannot be encrypted, so a password must not be requested.
     */
    public static final String BUILDER_EXPORT_CSV_PASSWORD_UNSUPPORTED = "builder.export.csv.passwordUnsupported";
    /**
     * The CSV charset is not a supported encoding name; {@code {0}}=sheet name, {@code {1}}=charset.
     */
    public static final String BUILDER_EXPORT_CSV_CHARSET_INVALID = "builder.export.csv.charsetInvalid";
    /**
     * The CSV field delimiter cannot be used; {@code {0}}=sheet name, {@code {1}}=delimiter.
     */
    public static final String BUILDER_EXPORT_CSV_DELIMITER_INVALID = "builder.export.csv.delimiterInvalid";
    /**
     * No import source was supplied.
     */
    public static final String BUILDER_IMPORT_SOURCE_MISSING = "builder.import.sourceMissing";
    /**
     * CSV import accepts a single source only.
     */
    public static final String BUILDER_IMPORT_CSV_SINGLE_SOURCE_ONLY = "builder.import.csv.singleSourceOnly";

    // Diagnostic log messages (SLF4J WARN, not thrown as exceptions). The *_FAILED keys take the
    // exception message (e.getMessage()) as their last positional parameter.

    /**
     * Bean-validation initialization failed (validation disabled); {@code {0}}=error.
     */
    public static final String LOG_BEAN_VALIDATION_INIT_FAILED = "log.beanValidationInitFailed";
    /**
     * Adding an image to the sheet failed; {@code {0}}=url, {@code {1}}=error.
     */
    public static final String LOG_IMAGE_ADD_FAILED = "log.imageAddFailed";
    /**
     * Applying a styler failed; {@code {0}}=stylerClass, {@code {1}}=error.
     */
    public static final String LOG_STYLER_APPLY_FAILED = "log.stylerApplyFailed";
    /**
     * Applying the quote-prefix style failed; {@code {0}}=error.
     */
    public static final String LOG_QUOTE_PREFIX_STYLE_FAILED = "log.quotePrefixStyleFailed";
    /**
     * Applying the date numeric style failed; {@code {0}}=error.
     */
    public static final String LOG_DATE_NUMERIC_STYLE_FAILED = "log.dateNumericStyleFailed";

}
