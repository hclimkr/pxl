package io.github.hclimkr.pxl.builder;

import io.github.hclimkr.pxl.exception.PxlException;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import io.github.hclimkr.pxl.util.PxlWorkbookUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * Common base for export/sample builders.
 *
 * <p>Subclass builders implement only {@link #build()}, which produces the workbook creation result ({@link Built}),
 * while the terminal methods (returning a workbook / file / stream) and resource handling and exception wrapping are shared here.</p>
 *
 * <p>Package-private: not part of the public API. Consumers reach the shared {@code public} terminals
 * ({@code toWorkbook()}/{@code toFile(...)}/{@code toStream(...)}) through the public concrete subclasses.</p>
 */
abstract class PxlAbstractExportBuilder {

    /**
     * Export option. Set via the subclass builder's {@code override(...)}. (Optional)
     */
    protected PxlExportWorkbookOption option;

    /**
     * Creates and returns a POI workbook. The returned workbook must be closed by the caller.
     *
     * @return the created workbook (the caller is responsible for closing it)
     * @throws PxlException if workbook creation fails
     */
    public final Workbook toWorkbook()
            throws PxlException {

        return build().workbook;
    }

    /**
     * Exports to an Excel file. (Encrypts if the option specifies a password)
     *
     * <p>The output stream to the file is opened and closed internally, so the caller has nothing to close.</p>
     *
     * @param excelFile the destination Excel file
     * @throws PxlException if workbook creation or writing fails
     */
    public final void toFile(final File excelFile)
            throws PxlException {

        PxlAssertSupport.notNull(excelFile, "excelFile");

        final Built built = build();
        OutputStream outputStream = null;

        try {
            outputStream = new BufferedOutputStream(new FileOutputStream(excelFile));
            PxlWorkbookUtils.writeToStream(built.workbook, outputStream, built.password);
        } catch (PxlException e) {
            throw e;
        } catch (Exception e) {
            throw new PxlException(e);
        } finally {
            IOUtils.closeQuietly(outputStream);
            PxlWorkbookUtils.closeWorkbook(built.workbook);
        }
    }

    /**
     * Exports to an Excel output stream. (Encrypts if the option specifies a password)
     *
     * <p>The given stream is flushed but <strong>not closed</strong>; the caller retains ownership and is responsible
     * for closing it.</p>
     *
     * @param outputStream the destination output stream (not closed by this method)
     * @throws PxlException if workbook creation or writing fails
     */
    public final void toStream(final OutputStream outputStream)
            throws PxlException {

        PxlAssertSupport.notNull(outputStream, "outputStream");

        final Built built = build();

        try {
            PxlWorkbookUtils.writeToStream(built.workbook, outputStream, built.password);
        } catch (PxlException e) {
            throw e;
        } catch (Exception e) {
            throw new PxlException(e);
        } finally {
            PxlWorkbookUtils.closeWorkbook(built.workbook);
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
