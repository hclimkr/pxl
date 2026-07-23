package io.github.hclimkr.pxl.internal.i18n;

import io.github.hclimkr.pxl.exception.PxlI18nException;

import java.util.ResourceBundle;

/**
 * Content i18n channel: translates consumer-facing <strong>sheet and column names</strong> through a
 * consumer-provided resource bundle, configured per workbook via {@code @PxlWorkbook}
 * ({@code importI18n*}/{@code exportI18n*}).
 *
 * <p>This is one of the library's two independent i18n channels. It is a thin, content-semantic facade over
 * the shared {@link PxlI18n} loader (UTF-8 bundle loading + key resolution); the other channel — the
 * library's own diagnostic/exception messages — is handled by {@link PxlI18nDiagnostic}. Keeping the content
 * channel separate from the loader makes the two concerns explicit: <em>consumer content translation</em>
 * (here) versus the <em>generic bundle mechanism</em> ({@link PxlI18n}) that both channels reuse.</p>
 *
 * <p>Like the loader, resolution is lenient: an absent bundle (i18n disabled) or a missing key yields the
 * original name unchanged, so translation gaps never fail a lookup.</p>
 */
public final class PxlI18nContent {

    /**
     * Prevents instantiation.
     */
    private PxlI18nContent() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Loads the consumer's sheet/column-name translation bundle for the given base name, language, and country.
     * <p>
     * Returns {@code null} when {@code baseName} is blank or when {@code language} or {@code country} is
     * {@code null}, treating this as "content i18n disabled" (see
     * {@link PxlI18n#getBundle(String, String, String)}). An empty (non-{@code null}) language or country is
     * treated as supplied and passed through.
     *
     * @param baseName the resource bundle base name; blank means content i18n is not used
     * @param language the ISO-639 language code; {@code null} means content i18n is not used
     * @param country  the ISO-3166 country code; {@code null} means content i18n is not used
     * @return the resolved bundle, or {@code null} when content i18n is disabled
     * @throws PxlI18nException if the bundle cannot be found for the given base name and locale
     */
    public static ResourceBundle loadBundle(final String baseName,
                                            final String language,
                                            final String country)
            throws PxlI18nException {

        return PxlI18n.getBundle(baseName, language, country);
    }

    /**
     * Translates a sheet or column {@code name} through the consumer {@code bundle}.
     * <p>
     * Returns {@code name} unchanged when the bundle is {@code null} (content i18n disabled) or the key is
     * absent, so untranslated names pass through verbatim.
     *
     * @param bundle the consumer bundle to read from; {@code null} yields the name unchanged
     * @param name   the sheet/column name, also used as the translation key and the fallback value
     * @return the translated name, or {@code name} when unresolved
     */
    public static String translate(final ResourceBundle bundle,
                                   final String name) {

        return PxlI18n.getMessage(bundle, name);
    }

}
