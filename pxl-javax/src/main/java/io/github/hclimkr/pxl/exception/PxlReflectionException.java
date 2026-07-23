package io.github.hclimkr.pxl.exception;

/**
 * Thrown when reflective access to a bound field/class fails (e.g. instantiation, field get/set, or
 * annotation resolution).
 */
public final class PxlReflectionException extends PxlException {

    /**
     * Creates an exception with no detail message.
     */
    public PxlReflectionException() {

        super();
    }

    /**
     * Creates an exception with the given detail message.
     *
     * @param message the detail message
     */
    public PxlReflectionException(final String message) {

        super(message);
    }

    /**
     * Creates an exception with the given detail message and cause.
     *
     * @param message the detail message
     * @param cause   the underlying cause
     */
    public PxlReflectionException(final String message, final Throwable cause) {

        super(message, cause);
    }

    /**
     * Creates an exception with the given cause.
     *
     * @param cause the underlying cause
     */
    public PxlReflectionException(final Throwable cause) {

        super(cause);
    }

}
