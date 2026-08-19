package io.github.hclimkr.pxl.util;

import com.github.pjfanning.xlsx.exceptions.OpenException;
import com.github.pjfanning.xlsx.exceptions.ReadException;
import com.github.pjfanning.xlsx.impl.StreamingWorkbook;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlIOException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.internal.support.PxlReflectionSupport;
import io.github.hclimkr.pxl.internal.support.PxlWorkbookSupport;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.UnsupportedFileFormatException;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.lang.reflect.Field;
import java.security.GeneralSecurityException;
import java.util.Objects;
import java.util.Optional;

/**
 * Workbook-related utilities: opening a workbook from a file or stream, writing it out (optionally encrypted),
 * closing it (disposing of a streaming workbook's temp files), creating a formula evaluator, and reading the
 * {@code @PxlWorkbookName} field of a workbook object.
 */
public final class PxlWorkbookUtils {

    /**
     * Prevents instantiation.
     */
    private PxlWorkbookUtils() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Opens an Excel workbook from a file, delegating format detection (XLS/XLSX) to POI's
     * {@link WorkbookFactory}. Every failure is translated into a {@link PxlIOException} carrying a
     * localized diagnostic message, so no non-Pxl exception escapes.
     *
     * @param workbookFile the workbook file to open
     * @param password     the password for an encrypted workbook, or {@code null} if none
     * @param readOnly     whether to open the workbook in read-only mode
     * @return the opened workbook
     * @throws PxlIOException          if the workbook cannot be opened - e.g. the password is wrong, the format
     *                                 is unsupported, the file does not exist, or it cannot be read
     * @throws PxlNullPointerException if {@code workbookFile} is {@code null}
     */
    public static Workbook openWorkbook(final File workbookFile,
                                        @Nullable final String password,
                                        final boolean readOnly)
            throws PxlIOException, PxlNullPointerException {

        PxlAssertSupport.notNull(workbookFile, "workbookFile");

        try {
            return WorkbookFactory.create(workbookFile, password, readOnly);
        } catch (EncryptedDocumentException e) {
            throw new PxlIOException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.UTIL_WORKBOOK_DECRYPT_FAILED), e);
        } catch (UnsupportedFileFormatException e) {
            // NOTE: the file may be DRM-protected.
            throw new PxlIOException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.UTIL_WORKBOOK_UNSUPPORTED_FORMAT), e);
        } catch (FileNotFoundException e) {
            throw new PxlIOException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.UTIL_WORKBOOK_FILE_NOT_FOUND), e);
        } catch (IOException | OpenException | ReadException e) {
            throw new PxlIOException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.UTIL_WORKBOOK_FILE_UNREADABLE), e);
        } catch (Exception e) {
            throw new PxlIOException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.UTIL_WORKBOOK_FILE_UNOPENABLE), e);
        }
    }

    /**
     * Opens an Excel workbook from an input stream, wrapping it in a {@link BufferedInputStream}
     * and delegating format detection (XLS/XLSX) to POI's {@link WorkbookFactory}. Every failure is
     * translated into a {@link PxlIOException} carrying a localized diagnostic message, so no non-Pxl
     * exception escapes.
     *
     * @param inputStream the stream to read the workbook from
     * @param password    the password for an encrypted workbook, or {@code null} if none
     * @return the opened workbook
     * @throws PxlIOException          if the workbook cannot be opened - e.g. the password is wrong, the format
     *                                 is unsupported, or the stream cannot be read
     * @throws PxlNullPointerException if {@code inputStream} is {@code null}
     */
    public static Workbook openWorkbook(final InputStream inputStream,
                                        @Nullable final String password)
            throws PxlIOException, PxlNullPointerException {

        PxlAssertSupport.notNull(inputStream, "inputStream");

        try {
            final BufferedInputStream bufferedInputStream = new BufferedInputStream(inputStream);

            return WorkbookFactory.create(bufferedInputStream, password);
        } catch (EncryptedDocumentException e) {
            throw new PxlIOException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.UTIL_WORKBOOK_DECRYPT_FAILED), e);
        } catch (UnsupportedFileFormatException e) {
            // NOTE: the file may be DRM-protected.
            throw new PxlIOException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.UTIL_WORKBOOK_UNSUPPORTED_FORMAT), e);
        } catch (IOException | ReadException e) {
            throw new PxlIOException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.UTIL_WORKBOOK_FILE_UNREADABLE), e);
        } catch (Exception e) {
            throw new PxlIOException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.UTIL_WORKBOOK_FILE_UNOPENABLE), e);
        }
    }

    /**
     * Writes the workbook to an output stream (export). When {@code password} is non-empty the output
     * is encrypted: XSSF/SXSSF workbooks are encrypted with agile OOXML encryption via a
     * {@link POIFSFileSystem}, while HSSF workbooks use Biff8 encryption (the thread-local key is
     * cleared afterwards). When {@code password} is empty or {@code null} the workbook is written as-is.
     * The stream is flushed but not closed.
     *
     * @param workbook     the workbook to write
     * @param outputStream the destination stream
     * @param password     the encryption password, or {@code null}/empty to write without encryption
     * @throws PxlIOException          if writing or encryption fails
     * @throws PxlArgumentException    if encryption is requested but the workbook type does not support it
     * @throws PxlNullPointerException if {@code workbook} or {@code outputStream} is {@code null}
     */
    public static void writeToStream(final Workbook workbook,
                                     final OutputStream outputStream,
                                     @Nullable final String password)
            throws PxlIOException, PxlArgumentException, PxlNullPointerException {

        PxlAssertSupport.notNull(workbook, "workbook");
        PxlAssertSupport.notNull(outputStream, "outputStream");

        try {
            if (StringUtils.isNotEmpty(password)) {
                if (workbook instanceof XSSFWorkbook || workbook instanceof SXSSFWorkbook) {

                    try (final POIFSFileSystem fileSystem = new POIFSFileSystem()) {
                        final EncryptionInfo encryptionInfo = new EncryptionInfo(EncryptionMode.agile);
                        final Encryptor encryptor = encryptionInfo.getEncryptor();
                        encryptor.confirmPassword(password);

                        try (final OutputStream os = encryptor.getDataStream(fileSystem)) {
                            // 1. first way
//                            final ByteArrayOutputStream baos = new ByteArrayOutputStream();
//                            workbook.write(baos);
//                            try (final OPCPackage opc = OPCPackage.open(new ByteArrayInputStream(baos.toByteArray()))) {
//                                opc.save(os);
//                            }

                            // 2. second way
                            workbook.write(os);
                        }

                        fileSystem.writeFilesystem(outputStream);
                    }
                } else if (workbook instanceof HSSFWorkbook) {
                    Biff8EncryptionKey.setCurrentUserPassword(password);
                    workbook.write(outputStream);
                } else {
                    throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.UTIL_WORKBOOK_ENCRYPT_UNSUPPORTED, workbook.getClass().getName()));
                }
            } else {
                workbook.write(outputStream);
            }

            outputStream.flush();
        } catch (IOException | GeneralSecurityException e) {
            throw new PxlIOException(e);
        } finally {
            if (StringUtils.isNotEmpty(password)) {
                if (workbook instanceof HSSFWorkbook) {
                    Biff8EncryptionKey.setCurrentUserPassword(null);
                }
            }
        }
    }

    /**
     * Closes the workbook, releasing its resources. For an {@link SXSSFWorkbook} the temporary files
     * backing it on disk are disposed first. A {@code null} workbook is a no-op, and any exception
     * raised while closing is swallowed.
     *
     * @param workbook the workbook to close, may be {@code null}
     */
    public static void closeWorkbook(final Workbook workbook) {

        if (Objects.isNull(workbook)) {
            return;
        }

        try {
            if (workbook instanceof SXSSFWorkbook) {
                // dispose of temporary files backing this workbook on disk
                ((SXSSFWorkbook) workbook).dispose();
            }
            workbook.close();
        } catch (Exception ignored) {
        }
    }

    /**
     * Creates a {@link FormulaEvaluator} for the workbook. A streaming workbook
     * ({@link StreamingWorkbook}) does not support formula evaluation, so {@code null} is returned;
     * {@code null} is also returned when the workbook exposes no creation helper.
     *
     * @param workbook the workbook to create an evaluator for
     * @return a formula evaluator, or {@code null} if the workbook is streaming or provides no creation helper
     * @throws PxlNullPointerException if {@code workbook} is {@code null}
     * @see <a href="https://poi.apache.org/components/spreadsheet/eval.html">POI formula evaluation</a>
     */
    public static FormulaEvaluator createFormulaEvaluator(final Workbook workbook)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(workbook, "workbook");

        // NOTE: StreamingWorkbook getCreationHelper() is not supported.
        final FormulaEvaluator formulaEvaluator = (workbook instanceof StreamingWorkbook)
                ? null
                : Optional.ofNullable(workbook.getCreationHelper())
                .map(CreationHelper::createFormulaEvaluator)
                .orElse(null);

        return formulaEvaluator;
    }

    /**
     * Finds and returns the workbook name carried by a workbook object's {@code @PxlWorkbookName} field.
     * <p>
     * This is a fail-safe lookup: it throws nothing. A {@code null} object, a class without a
     * {@code @PxlWorkbookName} field, and a field whose value cannot be read reflectively all yield {@code null}.
     *
     * @param workbookObject the workbook object to read the name from (may be {@code null})
     * @return the workbook name, or {@code null} if absent or unreadable
     */
    public static String getWorkbookNameFromWorkbookObject(@Nullable final Object workbookObject) {

        if (Objects.isNull(workbookObject)) {
            return null;
        }

        String workbookName = null;

        final Field workbookNameField = PxlWorkbookSupport.getWorkbookNameField(workbookObject.getClass());
        if (Objects.nonNull(workbookNameField)) {
            try {
                workbookName = (String) PxlReflectionSupport.getFieldValue(workbookNameField, workbookObject);
            } catch (Exception ignored) {
            }
        }

        return workbookName;
    }

}
