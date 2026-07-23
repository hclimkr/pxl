package io.github.hclimkr.pxl.exception;

/**
 * Thrown when a required (non-null) argument supplied to a builder/API call is {@code null}
 * (e.g. a required class, row collection, or source passed as {@code null}).
 */
public final class PxlNullPointerException extends PxlException {

    /**
     * Creates an exception with no detail message.
     */
    public PxlNullPointerException() {

        super();
    }

    /**
     * Creates an exception with the given detail message.
     *
     * @param message the detail message
     */
    public PxlNullPointerException(final String message) {

        super(message);
    }

    /**
     * Creates an exception with the given detail message and cause.
     *
     * @param message the detail message
     * @param cause   the underlying cause
     */
    public PxlNullPointerException(final String message, final Throwable cause) {

        super(message, cause);
    }

    /**
     * Creates an exception with the given cause.
     *
     * @param cause the underlying cause
     */
    public PxlNullPointerException(final Throwable cause) {

        super(cause);
    }

}
