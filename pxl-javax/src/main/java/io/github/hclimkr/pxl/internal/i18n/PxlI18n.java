package io.github.hclimkr.pxl.internal.i18n;

import io.github.hclimkr.pxl.exception.PxlI18nException;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;
import java.nio.charset.StandardCharsets;
import java.text.MessageFormat;
import java.util.*;

/**
 * Shared UTF-8 resource-bundle loader: the low-level i18n mechanism reused by <strong>both</strong>
 * of the library's i18n channels — the content channel ({@link PxlI18nContent}, sheet/column name
 * translation) and the diagnostic-message channel ({@link PxlI18nDiagnostic}, the library's own
 * exception/log text). It only loads bundles and resolves keys; the channels layer their own semantics on top.
 * <p>
 * Bundles are loaded through the nested {@link Utf8ResourceControl}, which decodes {@code .properties}
 * files as UTF-8 and disables JVM locale fallback. Missing keys and disabled i18n are handled
 * leniently so that lookups never fail on translation gaps.
 */
public class PxlI18n {

    private static final Utf8ResourceControl UTF8_RESOURCE_CONTROL = new Utf8ResourceControl();

    /**
     * Loads the resource bundle for the given base name and locale using the UTF-8 control.
     * <p>
     * Returns {@code null} when {@code baseName} is blank or {@code locale} is {@code null}, treating
     * this as "i18n disabled" to avoid accidental activation.
     *
     * @param baseName the resource bundle base name; blank means i18n is not used
     * @param locale   the locale to resolve; {@code null} means i18n is not used
     * @return the resolved bundle, or {@code null} when i18n is disabled
     * @throws PxlI18nException if the bundle cannot be found for the given base name and locale
     */
    public static ResourceBundle getBundle(final String baseName,
                                           final Locale locale)
            throws PxlI18nException {

        // If the base name is blank, treat i18n as unused and do not load the bundle. (Prevents accidental activation.)
        if (StringUtils.isBlank(baseName) || Objects.isNull(locale)) {
            return null;
        }

        try {
            return ResourceBundle.getBundle(baseName, locale, UTF8_RESOURCE_CONTROL);
        } catch (MissingResourceException e) {
            // Kept in English and NOT routed through PxlI18nDiagnostic on purpose: PxlI18nDiagnostic resolves the
            // library's own "pxl-messages" bundle through this very method, so localizing this failure would
            // recurse (and StackOverflow) whenever the pxl-messages bundle itself is missing. This is the i18n
            // loader reporting its own bootstrap failure, so it stays self-contained.
            throw new PxlI18nException("i18n resource bundle '" + baseName + "' (locale=" + locale + ") cannot be found. Check the i18nBaseName setting and the resource path.", e);
        }
    }

    /**
     * Loads the resource bundle for the given base name using the JVM default locale.
     *
     * @param baseName the resource bundle base name; blank means i18n is not used
     * @return the resolved bundle, or {@code null} when i18n is disabled
     * @throws PxlI18nException if the bundle cannot be found for the given base name
     */
    public static ResourceBundle getBundle(final String baseName)
            throws PxlI18nException {

        if (StringUtils.isBlank(baseName)) {
            return null;
        }

        return getBundle(baseName, Locale.getDefault());
    }

    /**
     * Loads the resource bundle for the given base name, language, and country.
     * <p>
     * Returns {@code null} when {@code baseName} is blank or when {@code language} or {@code country} is
     * {@code null}, treating this as "i18n disabled". An empty (non-{@code null}) language or country is
     * treated as supplied and passed through: the language/country pair is combined into a {@link Locale}
     * via {@link Locale#Locale(String, String)}.
     *
     * @param baseName the resource bundle base name; blank means i18n is not used
     * @param language the ISO-639 language code; {@code null} means i18n is not used
     * @param country  the ISO-3166 country code; {@code null} means i18n is not used
     * @return the resolved bundle, or {@code null} when i18n is disabled
     * @throws PxlI18nException if the bundle cannot be found for the given base name and locale
     */
    public static ResourceBundle getBundle(final String baseName,
                                           final String language,
                                           final String country)
            throws PxlI18nException {

        if (StringUtils.isBlank(baseName) || ObjectUtils.anyNull(language, country)) {
            return null;
        }

        return getBundle(baseName, new Locale(language, country));
    }

    /**
     * Loads the resource bundle for the given base name and IETF BCP 47 language tag.
     *
     * @param baseName    the resource bundle base name; blank means i18n is not used
     * @param languageTag the language tag parsed via {@link Locale#forLanguageTag(String)}; blank means i18n is not used
     * @return the resolved bundle, or {@code null} when i18n is disabled
     * @throws PxlI18nException if the bundle cannot be found for the given base name and locale
     */
    public static ResourceBundle getBundle(final String baseName,
                                           final String languageTag)
            throws PxlI18nException {

        if (StringUtils.isBlank(baseName) || StringUtils.isBlank(languageTag)) {
            return null;
        }

        return getBundle(baseName, Locale.forLanguageTag(languageTag));
    }

    /**
     * Resolves the message for the given key, optionally formatting it with the supplied parameters.
     * <p>
     * Returns the {@code key} unchanged when the bundle or key is {@code null}, when the key is blank,
     * or when the key is missing from the bundle. When {@code params} are provided, the resolved
     * message is formatted via {@link java.text.MessageFormat#format(String, Object...)}.
     *
     * @param resourceBundle the bundle to read from; {@code null} yields the key unchanged
     * @param key            the message key, also used as the fallback value
     * @param params         optional {@link java.text.MessageFormat} arguments
     * @return the resolved (and possibly formatted) message, or the key when unresolved
     */
    public static String getMessage(final ResourceBundle resourceBundle,
                                    final String key,
                                    final Object... params) {

        try {
            if (Objects.isNull(resourceBundle)) {
                return key;
            }

            if (StringUtils.isBlank(key)) {
                return key;
            }

            final String message = resourceBundle.getString(key);

            return ArrayUtils.isEmpty(params)
                    ? message
                    : MessageFormat.format(message, params);
        } catch (MissingResourceException e) {
            return key;
        }
    }

    /**
     * Resolves the message for the given key without parameter formatting.
     *
     * @param resourceBundle the bundle to read from; {@code null} yields the key unchanged
     * @param key            the message key, also used as the fallback value
     * @return the resolved message, or the key when unresolved
     */
    public static String getMessage(final ResourceBundle resourceBundle,
                                    final String key) {

        return getMessage(resourceBundle, key, (Object[]) null);
    }

    /**
     * Loads the bundle for {@code baseName}/{@code locale} and resolves {@code key} in a single step,
     * optionally formatting it with {@code params}.
     * <p>
     * Convenience over {@link #getBundle(String, Locale)} followed by
     * {@link #getMessage(ResourceBundle, String, Object...)}: the key is returned unchanged when i18n is
     * disabled (blank base name or {@code null} locale yields a {@code null} bundle) or when the key is missing.
     *
     * @param baseName the resource bundle base name; blank means i18n is not used
     * @param locale   the locale to resolve; {@code null} means i18n is not used
     * @param key      the message key, also used as the fallback value
     * @param params   optional {@link java.text.MessageFormat} arguments
     * @return the resolved (and possibly formatted) message, or the key when unresolved
     * @throws PxlI18nException if the bundle cannot be found for the given base name and locale
     */
    public static String getMessage(final String baseName,
                                    final Locale locale,
                                    final String key,
                                    final Object... params)
            throws PxlI18nException {

        final ResourceBundle resourceBundle = getBundle(baseName, locale);
        return getMessage(resourceBundle, key, params);
    }

    /**
     * Loads the bundle for {@code baseName}/{@code locale} and resolves {@code key} in a single step,
     * without parameter formatting.
     *
     * @param baseName the resource bundle base name; blank means i18n is not used
     * @param locale   the locale to resolve; {@code null} means i18n is not used
     * @param key      the message key, also used as the fallback value
     * @return the resolved message, or the key when unresolved
     * @throws PxlI18nException if the bundle cannot be found for the given base name and locale
     */
    public static String getMessage(final String baseName,
                                    final Locale locale,
                                    final String key)
            throws PxlI18nException {

        final ResourceBundle resourceBundle = getBundle(baseName, locale);
        return getMessage(resourceBundle, key);
    }

    /**
     * A {@link ResourceBundle.Control} that reads {@code .properties} bundles as UTF-8 and disables
     * JVM locale fallback.
     * <p>
     * Unlike the default control, which decodes properties with ISO-8859-1, this reader uses UTF-8 so
     * that non-Latin translations load correctly. It also restricts the format to {@code java.properties}
     * and returns {@code null} from {@link #getFallbackLocale} so unresolved locales do not silently fall
     * back to the JVM default locale bundle.
     * <p>
     * This is a pure implementation detail of {@link PxlI18n#getBundle(String, Locale)}; it is nested and
     * private because nothing outside this loader ever references the control directly.
     */
    private static final class Utf8ResourceControl extends ResourceBundle.Control {

        /**
         * Restricts bundle loading to the {@code java.properties} format only.
         *
         * @param baseName the resource bundle base name
         * @return a single-element list containing {@code "java.properties"}
         */
        @Override
        public List<String> getFormats(String baseName) {

            return Collections.singletonList("java.properties");
        }

        /**
         * Loads a {@code java.properties} bundle, decoding the resource stream as UTF-8.
         * <p>
         * Non-{@code java.properties} formats are delegated to the superclass. Returns {@code null} when
         * the resource does not exist.
         *
         * @param baseName the resource bundle base name
         * @param locale   the locale for which the bundle is loaded
         * @param format   the resource format; only {@code "java.properties"} is handled here
         * @param loader   the class loader used to locate the resource
         * @param reload   whether to bypass caches when opening the resource stream
         * @return the loaded bundle, or {@code null} if the resource is absent
         * @throws IllegalAccessException if delegated bundle instantiation is not accessible
         * @throws InstantiationException if delegated bundle instantiation fails
         * @throws IOException            if the resource stream cannot be read
         */
        @Override
        public ResourceBundle newBundle(String baseName,
                                        Locale locale,
                                        String format,
                                        ClassLoader loader,
                                        boolean reload)
                throws IllegalAccessException, InstantiationException, IOException {

            if (!StringUtils.equals("java.properties", format)) {
                return super.newBundle(baseName, locale, format, loader, reload);
            }

            final String bundleName = toBundleName(baseName, locale);
            final String resourceName = toResourceName(bundleName, "properties");

            try (InputStream is = getResourceStream(loader, reload, resourceName)) {
                if (Objects.isNull(is)) {
                    return null;
                }
                try (InputStreamReader reader = new InputStreamReader(is, StandardCharsets.UTF_8)) {
                    return new PropertyResourceBundle(reader);
                }
            }
        }

        /**
         * Disables JVM default-locale fallback by always returning {@code null}.
         *
         * @param baseName the resource bundle base name
         * @param locale   the locale that failed to resolve
         * @return {@code null}, indicating no fallback locale
         * @throws NullPointerException if {@code baseName} or {@code locale} is {@code null}
         */
        @Override
        public Locale getFallbackLocale(String baseName, Locale locale) {

            if (Objects.isNull(baseName) || Objects.isNull(locale)) {
                throw new NullPointerException();
            }

            return null;
        }

        /**
         * Opens an input stream for the named resource, optionally bypassing connection caches.
         *
         * @param loader       the class loader used to locate the resource
         * @param reload       whether to disable connection caching when opening the stream
         * @param resourceName the resource name to locate
         * @return the resource input stream, or {@code null} if the resource is absent
         * @throws IOException if the connection cannot be opened
         */
        private InputStream getResourceStream(final ClassLoader loader,
                                              final boolean reload,
                                              final String resourceName)
                throws IOException {

            final URL url = loader.getResource(resourceName);
            if (Objects.isNull(url)) {
                return null;
            }

            final URLConnection conn = url.openConnection();
            if (reload) {
                conn.setUseCaches(false);
            }

            return conn.getInputStream();
        }
    }

}
