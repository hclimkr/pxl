package io.github.hclimkr.pxl.exception;

/**
 * Thrown when a codec fails to convert between a cell value and the target field type — either parsing
 * a cell/string on import or building a cell on export.
 */
public final class PxlCellCodecException extends PxlException {

    /**
     * Creates an exception with no detail message.
     */
    public PxlCellCodecException() {

        super();
    }

    /**
     * Creates an exception with the given detail message.
     *
     * @param message the detail message
     */
    public PxlCellCodecException(final String message) {

        super(message);
    }

    /**
     * Creates an exception with the given detail message and cause.
     *
     * @param message the detail message
     * @param cause   the underlying cause
     */
    public PxlCellCodecException(final String message, final Throwable cause) {

        super(message, cause);
    }

    /**
     * Creates an exception with the given cause.
     *
     * @param cause the underlying cause
     */
    public PxlCellCodecException(final Throwable cause) {

        super(cause);
    }

    /**
     * Creates a location-tagged cell-codec exception.
     *
     * @param sheetName   sheet name (may be {@code null})
     * @param rowIndex    zero-based row index (may be {@code null})
     * @param columnName  column name (may be {@code null})
     * @param columnIndex zero-based column index, used only when {@code columnName} is {@code null} (may be {@code null})
     * @param message     the detail message appended after the location tag
     */
    public PxlCellCodecException(final String sheetName,
                                 final Integer rowIndex,
                                 final String columnName,
                                 final Integer columnIndex,
                                 final String message) {

        super(sheetName, rowIndex, columnName, columnIndex, message);
    }

    /**
     * Creates a location-tagged cell-codec exception with the underlying cause.
     *
     * @param sheetName   sheet name (may be {@code null})
     * @param rowIndex    zero-based row index (may be {@code null})
     * @param columnName  column name (may be {@code null})
     * @param columnIndex zero-based column index, used only when {@code columnName} is {@code null} (may be {@code null})
     * @param message     the detail message appended after the location tag
     * @param cause       the underlying cause
     */
    public PxlCellCodecException(final String sheetName,
                                 final Integer rowIndex,
                                 final String columnName,
                                 final Integer columnIndex,
                                 final String message,
                                 final Throwable cause) {

        super(sheetName, rowIndex, columnName, columnIndex, message, cause);
    }

    /**
     * Creates a location-tagged cell-codec exception, deriving the message from the cause.
     *
     * @param sheetName   sheet name (may be {@code null})
     * @param rowIndex    zero-based row index (may be {@code null})
     * @param columnName  column name (may be {@code null})
     * @param columnIndex zero-based column index, used only when {@code columnName} is {@code null} (may be {@code null})
     * @param cause       the underlying cause; its message is used as the detail message
     */
    public PxlCellCodecException(final String sheetName,
                                 final Integer rowIndex,
                                 final String columnName,
                                 final Integer columnIndex,
                                 final Throwable cause) {

        super(sheetName, rowIndex, columnName, columnIndex, cause);
    }

}
