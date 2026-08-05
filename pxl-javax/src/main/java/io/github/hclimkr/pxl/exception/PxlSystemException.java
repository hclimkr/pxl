package io.github.hclimkr.pxl.exception;

import java.io.IOException;

/**
 * Thrown when a failure that PXL does not classify reaches a builder's final (execute) step — the
 * normalization boundary at which anything that is not already a {@link PxlException} (a checked
 * {@link IOException}, an unexpected {@link RuntimeException} from POI/Commons CSV, ...) is wrapped so that
 * callers only ever face the {@link PxlException} family.
 *
 * <p>Since {@link PxlException} itself is abstract, this is the concrete fallback used when no more specific
 * {@code Pxl*Exception} applies; a failure PXL can classify is reported with that specific subtype instead
 * (e.g. {@link PxlIOException}, {@link PxlCellCodecException}, {@link PxlValidationException}). When it wraps
 * an underlying failure, that failure stays reachable through {@link Throwable#getCause()}.</p>
 */
public final class PxlSystemException extends PxlException {

    /**
     * Creates an exception with no detail message.
     */
    public PxlSystemException() {

        super();
    }

    /**
     * Creates an exception with the given detail message.
     *
     * @param message the detail message
     */
    public PxlSystemException(final String message) {

        super(message);
    }

    /**
     * Creates an exception with the given detail message and cause.
     *
     * @param message the detail message
     * @param cause   the underlying cause
     */
    public PxlSystemException(final String message, final Throwable cause) {

        super(message, cause);
    }

    /**
     * Creates an exception with the given cause. (The form used at the normalization boundary)
     *
     * @param cause the underlying cause
     */
    public PxlSystemException(final Throwable cause) {

        super(cause);
    }

}
