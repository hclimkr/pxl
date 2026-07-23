package io.github.hclimkr.pxl.exception;

import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.util.PxlMiscUtils;
import org.apache.commons.lang3.StringUtils;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * Base checked exception for the Pxl library and the common supertype of every {@code Pxl*Exception}.
 *
 * <p>All import/export failures are wrapped in this type (or one of its subtypes) at the {@code Pxl}
 * boundary. The context-aware constructors build a human-readable message prefixed with the sheet,
 * row, and column at which the error occurred (see {@link #buildTagMessage}).</p>
 */
public class PxlException extends Exception {

    /**
     * Creates an exception with no detail message.
     */
    public PxlException() {

        super();
    }

    /**
     * Creates an exception with the given detail message.
     *
     * @param message the detail message
     */
    public PxlException(final String message) {

        super(message);
    }

    /**
     * Creates an exception with the given detail message and cause.
     *
     * @param message the detail message
     * @param cause   the underlying cause
     */
    public PxlException(final String message, final Throwable cause) {

        super(message, cause);
    }

    /**
     * Creates an exception with the given cause.
     *
     * @param cause the underlying cause
     */
    public PxlException(final Throwable cause) {

        super(cause);
    }

    /**
     * Creates an exception whose detail message is tagged with the location (sheet/row/column) at which
     * the error occurred.
     *
     * @param sheetName   sheet name (may be {@code null} to omit)
     * @param rowIndex    zero-based row index, rendered one-based in the message (may be {@code null} to omit)
     * @param columnName  column name (may be {@code null})
     * @param columnIndex zero-based column index, used only when {@code columnName} is {@code null} (may be {@code null})
     * @param message     the detail message appended after the location tag
     */
    public PxlException(final String sheetName,
                        final Integer rowIndex,
                        final String columnName,
                        final Integer columnIndex,
                        final String message) {

        super(buildTagMessage(sheetName, rowIndex, columnName, columnIndex, message));
    }

    /**
     * Creates an exception whose detail message is tagged with the location (sheet/row/column) at which
     * the error occurred, together with the underlying cause.
     *
     * @param sheetName   sheet name (may be {@code null} to omit)
     * @param rowIndex    zero-based row index, rendered one-based in the message (may be {@code null} to omit)
     * @param columnName  column name (may be {@code null})
     * @param columnIndex zero-based column index, used only when {@code columnName} is {@code null} (may be {@code null})
     * @param message     the detail message appended after the location tag
     * @param cause       the underlying cause
     */
    public PxlException(final String sheetName,
                        final Integer rowIndex,
                        final String columnName,
                        final Integer columnIndex,
                        final String message,
                        final Throwable cause) {

        super(buildTagMessage(sheetName, rowIndex, columnName, columnIndex, message), cause);
    }

    /**
     * Creates an exception whose detail message is tagged with the location (sheet/row/column) at which
     * the error occurred, deriving the message from the cause's message.
     *
     * @param sheetName   sheet name (may be {@code null} to omit)
     * @param rowIndex    zero-based row index, rendered one-based in the message (may be {@code null} to omit)
     * @param columnName  column name (may be {@code null})
     * @param columnIndex zero-based column index, used only when {@code columnName} is {@code null} (may be {@code null})
     * @param cause       the underlying cause; its message is used as the detail message
     */
    public PxlException(final String sheetName,
                        final Integer rowIndex,
                        final String columnName,
                        final Integer columnIndex,
                        final Throwable cause) {

        super(buildTagMessage(sheetName, rowIndex, columnName, columnIndex,
                Objects.nonNull(cause) ? cause.getMessage() : null), cause);
    }

    /**
     * Builds a location-tagged message by joining the present sheet/row/column tags with commas and
     * prefixing them to {@code message} (as {@code "tag: message"}). When no location is present, the
     * message is returned unchanged. The row index is rendered one-based; when {@code columnName} is
     * absent, {@code columnIndex} is converted to a spreadsheet column string (A, B, ...).
     *
     * <p>The tag words and the join format are localized through the diagnostic message bundle
     * ({@code pxl-messages}): English by default, overridable process-wide via
     * {@link io.github.hclimkr.pxl.Pxl#setMessageLocale(java.util.Locale)}.</p>
     *
     * @param sheetName   sheet name (may be {@code null})
     * @param rowIndex    zero-based row index (may be {@code null})
     * @param columnName  column name (may be {@code null})
     * @param columnIndex zero-based column index, used only when {@code columnName} is {@code null} (may be {@code null})
     * @param message     the message to append after the location tag
     * @return the location-tagged message
     */
    protected static String buildTagMessage(final String sheetName,
                                            final Integer rowIndex,
                                            final String columnName,
                                            final Integer columnIndex,
                                            final String message) {

        final List<String> tags = new ArrayList<>();

        if (Objects.nonNull(sheetName)) {
            tags.add(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.TAG_SHEET, sheetName));
        }

        if (Objects.nonNull(rowIndex)) {
            tags.add(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.TAG_ROW, String.valueOf(rowIndex + 1)));
        }

        if (Objects.nonNull(columnName)) {
            tags.add(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.TAG_COLUMN_NAME, columnName));
        } else if (Objects.nonNull(columnIndex)) {
            tags.add(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.TAG_COLUMN_INDEX,
                    PxlMiscUtils.convertColumnIndexToColumnString(columnIndex)));
        }

        final String tag = StringUtils.join(tags, ", ");
        if (StringUtils.isNotEmpty(tag)) {
            return PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.TAG_JOIN, tag, message);
        } else {
            return message;
        }
    }

}
