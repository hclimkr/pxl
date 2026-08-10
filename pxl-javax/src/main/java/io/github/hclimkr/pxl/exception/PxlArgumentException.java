package io.github.hclimkr.pxl.exception;

/**
 * Thrown when an argument or a declared configuration value is present but invalid or inconsistent - a
 * builder/API argument that is empty, blank, or contradictory (missing or mutually exclusive builder inputs,
 * an invalid cell reference), or an annotation/option value that cannot be applied (an unparseable
 * date-time/number pattern, a grouping field that does not exist on the row class, a misdeclared custom
 * converter, an unsupported field type).
 *
 * <p>An operation the current configuration cannot support is reported the same way - e.g. reading merged
 * regions through the streaming reader, or encrypting a workbook type that does not support it. A required
 * argument that is {@code null} raises {@link PxlNullPointerException} instead.</p>
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
