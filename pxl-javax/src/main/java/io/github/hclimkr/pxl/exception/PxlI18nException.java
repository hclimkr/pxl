package io.github.hclimkr.pxl.exception;

/**
 * Thrown when an i18n resource bundle cannot be found for a configured base name and locale — the bundle
 * is absent from the resource path. Raised by the shared i18n loader while opening either a consumer
 * content bundle (sheet/column name translation) or the library's own {@code pxl-messages} diagnostic
 * bundle. A blank base name or {@code null} locale is treated as "i18n disabled" and does <em>not</em>
 * raise this exception.
 */
public final class PxlI18nException extends PxlException {

    /**
     * Creates an exception with no detail message.
     */
    public PxlI18nException() {

        super();
    }

    /**
     * Creates an exception with the given detail message.
     *
     * @param message the detail message
     */
    public PxlI18nException(final String message) {

        super(message);
    }

    /**
     * Creates an exception with the given detail message and cause.
     *
     * @param message the detail message
     * @param cause   the underlying cause
     */
    public PxlI18nException(final String message, final Throwable cause) {

        super(message, cause);
    }

    /**
     * Creates an exception with the given cause.
     *
     * @param cause the underlying cause
     */
    public PxlI18nException(final Throwable cause) {

        super(cause);
    }

}
