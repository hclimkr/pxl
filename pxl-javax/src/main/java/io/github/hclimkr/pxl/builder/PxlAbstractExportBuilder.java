package io.github.hclimkr.pxl.builder;

import io.github.hclimkr.pxl.exception.*;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import org.apache.commons.io.IOUtils;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * Format-neutral common base for the export builders.
 *
 * <p>Holds the export option and the file/stream terminal methods, which own nothing but resource handling
 * and exception normalization. What is actually written is left to three seams the subclass fills in:</p>
 * <ol>
 *   <li>{@link #prepare()} - runs <strong>before</strong> the destination is opened, so a failure here leaves
 *       no file behind</li>
 *   <li>{@link #writeTo(OutputStream)} - writes the result to the destination</li>
 *   <li>{@link #cleanup()} - releases whatever {@code prepare()} acquired, exactly once, on success and failure alike</li>
 * </ol>
 *
 * <p>The order matters: preparing after the destination was opened would leave an empty file whenever the
 * preparation fails, and releasing outside the {@code finally} would leak whatever was prepared whenever
 * opening the destination fails. All three run inside one {@code try}/{@code finally}, so the guarantees hold
 * from the first seam onward - a failure in {@code prepare()} is normalized like any other and still reaches
 * {@code cleanup()}, rather than escaping raw because of where it happened to be raised.</p>
 *
 * <p>The terminal methods are the <strong>normalization boundary</strong>: they declare
 * {@code throws PxlException}, but since that type is abstract what actually surfaces is always a concrete
 * subtype - the matching one for a classified failure ({@link PxlArgumentException},
 * {@link PxlCellCodecException},
 * {@link PxlValidationException}, ...), and
 * {@link PxlSystemException} (carrying the original as its cause) for anything else, including checked I/O failures
 * and unexpected runtime failures from either seam. An {@link Error} is not covered: it is not an
 * {@link Exception}, so an {@link OutOfMemoryError} raised while a result is being prepared surfaces as itself.
 * Catching it would be wrong, since wrapping allocates in the very condition that ran out of memory.</p>
 *
 * <p>Package-private: not part of the public API. Consumers reach the shared {@code public} terminals
 * ({@code toFile(...)}/{@code toStream(...)}) through the public concrete subclasses.</p>
 */
abstract class PxlAbstractExportBuilder {

    /**
     * Export option. Set via the subclass builder's {@code override(...)}. (Optional)
     */
    protected PxlExportWorkbookOption option;

    /**
     * Exports to a file. (Encrypts if the option specifies a password and the format supports it)
     *
     * <p>The output stream to the file is opened and closed internally, so the caller has nothing to close.</p>
     *
     * @param outputFile the destination file
     * @throws PxlNullPointerException if {@code outputFile} is {@code null}
     * @throws PxlException            if preparing the result or writing fails
     */
    public final void toFile(final File outputFile)
            throws PxlException {

        PxlAssertSupport.notNull(outputFile, "outputFile");

        OutputStream outputStream = null;

        try {
            // Prepare before the destination is opened: a failure here must not leave an empty file behind.
            prepare();

            outputStream = new BufferedOutputStream(new FileOutputStream(outputFile));
            writeTo(outputStream);
        } catch (PxlException e) {
            throw e;
        } catch (Exception e) {
            throw new PxlSystemException(e);
        } finally {
            IOUtils.closeQuietly(outputStream);
            // Opening the destination may have failed, so releasing has to happen here rather than in writeTo.
            cleanup();
        }
    }

    /**
     * Exports to an output stream. (Encrypts if the option specifies a password and the format supports it)
     *
     * <p>The given stream is flushed but <strong>not closed</strong>; the caller retains ownership and is responsible
     * for closing it.</p>
     *
     * @param outputStream the destination output stream (not closed by this method)
     * @throws PxlNullPointerException if {@code outputStream} is {@code null}
     * @throws PxlException            if preparing the result or writing fails
     */
    public final void toStream(final OutputStream outputStream)
            throws PxlException {

        PxlAssertSupport.notNull(outputStream, "outputStream");

        try {
            prepare();

            writeTo(outputStream);
        } catch (PxlException e) {
            throw e;
        } catch (Exception e) {
            throw new PxlSystemException(e);
        } finally {
            cleanup();
        }
    }

    /**
     * Prepares the result to be written, before the destination is opened.
     *
     * <p>Implementations must prepare afresh on every call - a terminal method run twice repeats the work
     * rather than handing out a cached result. A format with nothing to prepare implements this as an empty
     * body; it is left abstract so that every format has to state its preparation explicitly rather than
     * inherit silence (a skipped validation would otherwise pass unnoticed).</p>
     *
     * @throws PxlException if the result cannot be prepared
     */
    protected abstract void prepare()
            throws PxlException;

    /**
     * Writes the prepared result to the given destination. The stream is owned by the caller of this method
     * (the terminal), so implementations must not close it.
     *
     * @param outputStream the destination output stream
     * @throws PxlException if writing fails
     */
    protected abstract void writeTo(OutputStream outputStream)
            throws PxlException;

    /**
     * Releases whatever {@link #prepare()} acquired. Called exactly once from the terminal's {@code finally},
     * so it runs on success and failure alike, and must not throw.
     *
     * <p>A format that acquires nothing implements this as an empty body; it is left abstract because it pairs
     * with {@link #prepare()} - whoever acquires a resource there has to say here how it is released.</p>
     */
    protected abstract void cleanup();

}
