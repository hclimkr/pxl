package io.github.hclimkr.pxl.exception;

/**
 * Thrown when the source data violates a structural limit or expectation (e.g. exceeding the maximum
 * number of sheets/rows/columns, or an otherwise malformed workbook/CSV shape).
 */
public final class PxlDataException extends PxlException {

    /**
     * Creates an exception with no detail message.
     */
    public PxlDataException() {

        super();
    }

    /**
     * Creates an exception with the given detail message.
     *
     * @param message the detail message
     */
    public PxlDataException(final String message) {

        super(message);
    }

    /**
     * Creates an exception with the given detail message and cause.
     *
     * @param message the detail message
     * @param cause   the underlying cause
     */
    public PxlDataException(final String message, final Throwable cause) {

        super(message, cause);
    }

    /**
     * Creates an exception with the given cause.
     *
     * @param cause the underlying cause
     */
    public PxlDataException(final Throwable cause) {

        super(cause);
    }

}
