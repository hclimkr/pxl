package io.github.hclimkr.pxl.exception;

/**
 * Thrown when reading a source (file/stream) or writing a workbook to a file/stream fails at the I/O level.
 */
public final class PxlIOException extends PxlException {

    /**
     * Creates an exception with no detail message.
     */
    public PxlIOException() {

        super();
    }

    /**
     * Creates an exception with the given detail message.
     *
     * @param message the detail message
     */
    public PxlIOException(final String message) {

        super(message);
    }

    /**
     * Creates an exception with the given detail message and cause.
     *
     * @param message the detail message
     * @param cause   the underlying cause
     */
    public PxlIOException(final String message, final Throwable cause) {

        super(message, cause);
    }

    /**
     * Creates an exception with the given cause.
     *
     * @param cause the underlying cause
     */
    public PxlIOException(final Throwable cause) {

        super(cause);
    }

}
