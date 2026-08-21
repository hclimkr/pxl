package io.github.hclimkr.pxl.exception;

/**
 * Thrown when bean-validation constraints on a bound row/object are violated during export or import.
 */
public final class PxlValidationException extends PxlException {

    /**
     * Creates an exception with no detail message.
     */
    public PxlValidationException() {

        super();
    }

    /**
     * Creates an exception with the given detail message.
     *
     * @param message the detail message
     */
    public PxlValidationException(final String message) {

        super(message);
    }

    /**
     * Creates an exception with the given detail message and cause.
     *
     * @param message the detail message
     * @param cause   the underlying cause
     */
    public PxlValidationException(final String message, final Throwable cause) {

        super(message, cause);
    }

    /**
     * Creates an exception with the given cause.
     *
     * @param cause the underlying cause
     */
    public PxlValidationException(final Throwable cause) {

        super(cause);
    }

    /**
     * Creates a location-tagged validation exception.
     *
     * @param sheetName   sheet name (may be {@code null})
     * @param rowIndex    zero-based row index (may be {@code null})
     * @param columnName  column name (may be {@code null})
     * @param columnIndex zero-based column index, used only when {@code columnName} is {@code null} (may be {@code null})
     * @param message     the detail message appended after the location tag
     */
    public PxlValidationException(final String sheetName,
                                  final Integer rowIndex,
                                  final String columnName,
                                  final Integer columnIndex,
                                  final String message) {

        super(sheetName, rowIndex, columnName, columnIndex, message);
    }

    /**
     * Creates a location-tagged validation exception with the underlying cause.
     *
     * @param sheetName   sheet name (may be {@code null})
     * @param rowIndex    zero-based row index (may be {@code null})
     * @param columnName  column name (may be {@code null})
     * @param columnIndex zero-based column index, used only when {@code columnName} is {@code null} (may be {@code null})
     * @param message     the detail message appended after the location tag
     * @param cause       the underlying cause
     */
    public PxlValidationException(final String sheetName,
                                  final Integer rowIndex,
                                  final String columnName,
                                  final Integer columnIndex,
                                  final String message,
                                  final Throwable cause) {

        super(sheetName, rowIndex, columnName, columnIndex, message, cause);
    }

    /**
     * Creates a location-tagged validation exception, deriving the message from the cause.
     *
     * @param sheetName   sheet name (may be {@code null})
     * @param rowIndex    zero-based row index (may be {@code null})
     * @param columnName  column name (may be {@code null})
     * @param columnIndex zero-based column index, used only when {@code columnName} is {@code null} (may be {@code null})
     * @param cause       the underlying cause; its message is used as the detail message
     */
    public PxlValidationException(final String sheetName,
                                  final Integer rowIndex,
                                  final String columnName,
                                  final Integer columnIndex,
                                  final Throwable cause) {

        super(sheetName, rowIndex, columnName, columnIndex, cause);
    }

}
