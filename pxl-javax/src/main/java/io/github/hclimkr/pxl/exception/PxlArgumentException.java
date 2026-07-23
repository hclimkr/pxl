package io.github.hclimkr.pxl.exception;

/**
 * Thrown when an invalid or inconsistent configuration argument is supplied to a builder/API call
 * (e.g. missing or mutually exclusive builder inputs, illegal option values).
 */
public final class PxlArgumentException extends PxlException {

    /**
     * Creates an exception with no detail message.
     */
    public PxlArgumentException() {

        super();
    }

    /**
     * Creates an exception with the given detail message.
     *
     * @param message the detail message
     */
    public PxlArgumentException(final String message) {

        super(message);
    }

    /**
     * Creates an exception with the given detail message and cause.
     *
     * @param message the detail message
     * @param cause   the underlying cause
     */
    public PxlArgumentException(final String message, final Throwable cause) {

        super(message, cause);
    }

    /**
     * Creates an exception with the given cause.
     *
     * @param cause the underlying cause
     */
    public PxlArgumentException(final Throwable cause) {

        super(cause);
    }

}
