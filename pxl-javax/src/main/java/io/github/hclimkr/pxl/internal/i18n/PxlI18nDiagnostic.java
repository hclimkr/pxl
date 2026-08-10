package io.github.hclimkr.pxl.internal.i18n;

import java.text.MessageFormat;
import java.util.Locale;
import java.util.Objects;
import java.util.ResourceBundle;

/**
 * Resolves the PXL library's own diagnostic messages (exception text and the location tag) from the
 * {@code pxl-messages} bundle shipped inside the pxl artifact.
 *
 * <p><strong>Separate from consumer content i18n.</strong> Sheet/column name translation (the content
 * channel, {@link PxlI18nContent}) is driven per workbook by {@code @PxlWorkbook} (a consumer-provided
 * bundle and locale). Diagnostic messages are a
 * library-owned, <em>process-wide</em> concern instead: the locale is the JVM default
 * ({@link Locale#getDefault()}), optionally overridden globally via {@link #setOverrideLocale(Locale)}
 * (exposed on the public entry point as {@code Pxl.setMessageLocale}). The base bundle
 * ({@code pxl-messages.properties}) is English; {@code pxl-messages_ko.properties} adds Korean, and any
 * unmatched locale falls back to the English base.</p>
 *
 * <p>Resolution is <strong>fail-safe</strong>: {@link #get(String, Object...)} never throws, so building an
 * exception message can never itself fail and mask the original error - on any lookup failure it returns the
 * key unchanged (matching {@link PxlI18n#getMessage}'s lenient behavior for missing keys).</p>
 */
public final class PxlI18nDiagnostic {

    /**
     * Base name of the library's diagnostic bundle. Namespaced to avoid clashing with a consumer's own {@code messages} bundle.
     */
    static final String BASE_NAME = "pxl-messages";

    private static volatile Locale overrideLocale = null;

    /**
     * Prevents instantiation.
     */
    private PxlI18nDiagnostic() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Overrides the locale used for all diagnostic messages process-wide.
     *
     * @param locale the locale to use; {@code null} clears the override and reverts to {@link Locale#getDefault()}
     */
    public static void setOverrideLocale(final Locale locale) {

        overrideLocale = locale;
    }

    /**
     * Returns the current locale override, or {@code null} when none is set (JVM default in effect).
     *
     * @return the override locale, or {@code null}
     */
    public static Locale getOverrideLocale() {

        return overrideLocale;
    }

    /**
     * Resolves the current diagnostic locale: the override when set, otherwise the JVM default.
     *
     * @return the effective locale for diagnostic messages
     */
    static Locale currentLocale() {

        final Locale locale = overrideLocale;

        return Objects.nonNull(locale) ? locale : Locale.getDefault();
    }

    /**
     * Resolves and formats the diagnostic message for the given key.
     * <p>
     * Never throws: on any failure (bundle missing, key missing, formatting error) the {@code key} is returned
     * unchanged so that message resolution cannot mask the error it was describing.
     *
     * @param key    the message key from {@link PxlI18nDiagnosticKeys}
     * @param params optional {@link MessageFormat} arguments
     * @return the resolved (and possibly formatted) message, or {@code key} when unresolved
     */
    public static String get(final String key, final Object... params) {

        try {
            final ResourceBundle bundle = PxlI18n.getBundle(BASE_NAME, currentLocale());

            return PxlI18n.getMessage(bundle, key, params);
        } catch (Exception e) {
            return key;
        }
    }

}
