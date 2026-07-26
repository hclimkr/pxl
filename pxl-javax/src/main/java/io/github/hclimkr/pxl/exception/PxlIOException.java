package io.github.hclimkr.pxl.exception;

/**
 * Thrown when a source (file/stream) cannot be opened or read, or when writing a workbook to a file/stream fails.
 *
 * <p>Besides plain I/O failures, the container-level reasons that keep a source from being opened map here as
 * well: a file that does not exist, a file that cannot be opened or read, an unsupported spreadsheet format,
 * and a password-protected workbook that cannot be decrypted (a missing or wrong {@code importPassword}). The
 * underlying failure is kept as the cause.</p>
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
