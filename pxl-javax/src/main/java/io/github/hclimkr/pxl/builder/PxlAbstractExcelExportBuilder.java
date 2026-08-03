package io.github.hclimkr.pxl.builder;

import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlException;
import io.github.hclimkr.pxl.exception.PxlIOException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.util.PxlWorkbookUtils;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.OutputStream;
import java.util.Objects;

/**
 * Common base for the Excel (POI workbook) export builders.
 *
 * <p>Adds everything POI-specific on top of {@link PxlAbstractExportBuilder}: the workbook terminal
 * ({@link #toWorkbook()}), the creation result holder ({@link Built}), and the three write seams implemented
 * in terms of a POI workbook. Subclass builders implement only {@link #build()}.</p>
 *
 * <p>The workbook is created in {@link #prepare()} — that is, before the destination is opened — and released in
 * {@link #cleanup()}, so a workbook is never left open when opening the destination fails. It is built afresh on
 * every terminal call: running a terminal twice repeats the build rather than reusing the previous workbook.</p>
 *
 * <p>Package-private: not part of the public API. Consumers reach the shared {@code public} terminals
 * ({@code toWorkbook()}/{@code toFile(...)}/{@code toStream(...)}) through the public concrete subclasses.</p>
 */
abstract class PxlAbstractExcelExportBuilder extends PxlAbstractExportBuilder {

    /**
     * Result of the current terminal call, created by {@link #prepare()} and released by {@link #cleanup()}.
     */
    private Built built;

    /**
     * Creates and returns a POI workbook. The returned workbook must be closed by the caller.
     *
     * <p>The resolved {@code exportPassword} is <strong>not</strong> applied to the returned workbook: POI cannot carry
     * a document-open password on the workbook object itself (encryption happens at the file-container layer), so PXL
     * encrypts only while writing. Writing the returned workbook with {@code Workbook.write(...)} therefore produces an
     * unencrypted file — write it with {@link PxlWorkbookUtils#writeToStream(Workbook, OutputStream, String)} to have
     * the password applied, or use {@link #toFile(File)} / {@link #toStream(OutputStream)} instead.</p>
     *
     * @return the created workbook (the caller is responsible for closing it)
     * @throws PxlException if workbook creation fails
     */
    public final Workbook toWorkbook()
            throws PxlException {

        return build().workbook;
    }

    /**
     * Builds the workbook for the current terminal call.
     *
     * @throws PxlException if workbook creation fails
     */
    @Override
    protected final void prepare()
            throws PxlException {

        this.built = build();
    }

    /**
     * Writes the built workbook to the given destination, encrypting it when a password was resolved.
     *
     * @param outputStream the destination output stream (not closed here)
     * @throws PxlIOException          if writing the workbook fails
     * @throws PxlArgumentException    if the encryption password cannot be applied
     * @throws PxlNullPointerException if the built workbook is missing
     */
    @Override
    protected final void writeTo(final OutputStream outputStream)
            throws PxlIOException, PxlArgumentException, PxlNullPointerException {

        PxlWorkbookUtils.writeToStream(built.workbook, outputStream, built.password);
    }

    /**
     * Closes the built workbook (disposing of any streaming temp files) and drops it, so the next terminal call
     * builds a fresh one.
     */
    @Override
    protected final void cleanup() {

        if (Objects.nonNull(built)) {
            PxlWorkbookUtils.closeWorkbook(built.workbook);
            this.built = null;
        }
    }

    /**
     * Creates the workbook and (optional) password. Implemented by the subclass builder for data/sample.
     *
     * @return the creation result
     * @throws PxlException if creation fails
     */
    protected abstract Built build()
            throws PxlException;

    /**
     * The {@link #build()} result — the created workbook and (optional) password.
     */
    protected static final class Built {

        protected final Workbook workbook;
        protected final String password;

        /**
         * Holds the created workbook and (optional) password.
         *
         * @param workbook the created workbook
         * @param password the encryption password, or {@code null}/empty for no encryption
         */
        protected Built(final Workbook workbook, final String password) {

            this.workbook = workbook;
            this.password = password;
        }

    }

}
