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
 * bundle and locale). Diagnostic messages are a library-owned concern instead, with their own locale.
 * The base bundle ({@code pxl-messages.properties}) is English; {@code pxl-messages_ko.properties} adds
 * Korean, and any unmatched locale falls back to the English base.</p>
 *
 * <p><strong>Two-tier locale.</strong> {@link #currentLocale()} takes the first of these that is set,
 * narrowest first:</p>
 * <ol>
 *   <li>the calling thread's override ({@link #setThreadOverrideLocale(Locale)}, exposed on the public
 *       entry point as {@code Pxl.setThreadMessageLocale}) - for a request-scoped language, e.g. a
 *       servlet filter mirroring the container's or framework's per-request locale;</li>
 *   <li>the process-wide override ({@link #setGlobalOverrideLocale(Locale)}, exposed as
 *       {@code Pxl.setMessageLocale}) - the "set it once at startup" form, visible to every thread;</li>
 *   <li>the JVM default ({@link Locale#getDefault()}).</li>
 * </ol>
 *
 * <p><strong>Caller responsibility for the thread tier.</strong> Because a thread override lives on the
 * calling thread, it must be cleared (via {@link #setThreadOverrideLocale(Locale)} with {@code null}, or
 * {@code Pxl.resetThreadMessageLocale()}) before the thread returns to a pool - otherwise a later,
 * unrelated task scheduled on that same thread observes the stale override. A thread override also does
 * not follow work across thread boundaries (thread pool hand-off, {@code CompletableFuture.supplyAsync},
 * reactive schedulers); set it again on the new thread if it still applies there. The process-wide tier
 * has neither concern - it is a single {@code volatile} field every thread reads.</p>
 *
 * <p>Resolution is <strong>fail-safe</strong>: {@link #get(String, Object...)} never throws, so building an
 * exception message can never itself fail and mask the original error - on any lookup failure it returns the
 * key unchanged (matching {@link PxlI18n#getMessage}'s lenient behavior for missing keys).</p>
 */
public final class PxlI18nDiagnostic {

    /**
     * Base name of the library's diagnostic bundle. Namespaced to avoid clashing with a consumer's own {@code messages} bundle.
     */
    private static final String BASE_NAME = "pxl-messages";

    private static volatile Locale globalOverrideLocale = null;

    private static final ThreadLocal<Locale> THREAD_OVERRIDE_LOCALE = new ThreadLocal<>();

    /**
     * Prevents instantiation.
     */
    private PxlI18nDiagnostic() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Overrides the locale used for diagnostic messages on the calling thread only, taking precedence
     * over the process-wide override set by {@link #setGlobalOverrideLocale(Locale)}.
     *
     * <p>The override is invisible to other threads and does not follow the calling thread's work if
     * that work later resumes on a different thread. Callers running on a pooled thread (e.g. a servlet
     * container's request-handling thread) must clear the override before the thread returns to the
     * pool, or a later unrelated task on that thread will observe it.</p>
     *
     * @param locale the locale to use on this thread; {@code null} clears the thread override, so this
     *               thread falls back to the process-wide override and then to {@link Locale#getDefault()}
     */
    public static void setThreadOverrideLocale(final Locale locale) {

        if (Objects.isNull(locale)) {
            THREAD_OVERRIDE_LOCALE.remove();
        } else {
            THREAD_OVERRIDE_LOCALE.set(locale);
        }
    }

    /**
     * Returns the calling thread's locale override, or {@code null} when none is set on this thread
     * (the process-wide override or the JVM default is then in effect).
     *
     * @return the thread override locale, or {@code null}
     */
    public static Locale getThreadOverrideLocale() {

        return THREAD_OVERRIDE_LOCALE.get();
    }

    /**
     * Overrides the locale used for diagnostic messages on every thread that has no thread override of
     * its own. Set once at startup, it stays in effect for the life of the process.
     *
     * @param locale the locale to use process-wide; {@code null} clears the override and reverts to
     *               {@link Locale#getDefault()}
     */
    public static void setGlobalOverrideLocale(final Locale locale) {

        globalOverrideLocale = locale;
    }

    /**
     * Returns the process-wide locale override, or {@code null} when none is set (JVM default in effect
     * for threads without an override of their own).
     *
     * @return the process-wide override locale, or {@code null}
     */
    public static Locale getGlobalOverrideLocale() {

        return globalOverrideLocale;
    }

    /**
     * Resolves the effective diagnostic locale for the calling thread: its thread override when set,
     * otherwise the process-wide override, otherwise the JVM default.
     *
     * @return the effective locale for diagnostic messages on this thread
     */
    static Locale currentLocale() {

        final Locale threadLocale = THREAD_OVERRIDE_LOCALE.get();
        if (Objects.nonNull(threadLocale)) {
            return threadLocale;
        }

        if (Objects.nonNull(globalOverrideLocale)) {
            return globalOverrideLocale;
        }

        return Locale.getDefault();
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
