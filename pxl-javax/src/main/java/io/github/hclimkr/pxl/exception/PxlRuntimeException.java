package io.github.hclimkr.pxl.exception;

/**
 * Unchecked counterpart to {@link PxlException} (it extends {@link RuntimeException}, so it is <em>not</em> caught by
 * {@code catch (PxlException ...)}), intended for unrecoverable initialization/configuration failures at a public
 * boundary where a checked exception cannot be declared, keeping such failures within a Pxl-namespaced type instead of
 * leaking a raw third-party exception (e.g. {@code javax.validation.ValidationException}).
 *
 * <p><strong>Currently unused.</strong> Its only former use was the {@link io.github.hclimkr.pxl.Pxl} constructor when
 * bean-validation bootstrap failed; that path now graceful-degrades instead (disables validation and logs a warning via
 * SLF4J), so this type is retained but thrown nowhere in the current source.</p>
 */
public final class PxlRuntimeException extends RuntimeException {

    /**
     * Creates an exception with no detail message.
     */
    public PxlRuntimeException() {

        super();
    }

    /**
     * Creates an exception with the given detail message.
     *
     * @param message the detail message
     */
    public PxlRuntimeException(final String message) {

        super(message);
    }

    /**
     * Creates an exception with the given detail message and cause.
     *
     * @param message the detail message
     * @param cause   the underlying cause
     */
    public PxlRuntimeException(final String message, final Throwable cause) {

        super(message, cause);
    }

    /**
     * Creates an exception with the given cause.
     *
     * @param cause the underlying cause
     */
    public PxlRuntimeException(final Throwable cause) {

        super(cause);
    }

}
