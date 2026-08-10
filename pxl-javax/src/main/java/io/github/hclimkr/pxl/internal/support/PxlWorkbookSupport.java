package io.github.hclimkr.pxl.internal.support;

import com.github.pjfanning.xlsx.StreamingReader;
import com.github.pjfanning.xlsx.exceptions.OpenException;
import com.github.pjfanning.xlsx.exceptions.ReadException;
import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import io.github.hclimkr.pxl.exception.PxlDataException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlImportWorkbookMeta;
import io.github.hclimkr.pxl.type.PxlExcelEngine;
import io.github.hclimkr.pxl.type.PxlFileFormat;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import io.github.hclimkr.pxl.util.PxlWorkbookUtils;
import org.apache.commons.io.input.CloseShieldInputStream;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.UnsupportedFileFormatException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.lang.reflect.Field;
import java.nio.file.InvalidPathException;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Workbook lifecycle and naming for the binder: creating an empty workbook for an export engine, opening a source
 * workbook for import, resolving the {@code @PxlWorkbookName} field, and making sheet/defined names unique.
 * <p>
 * {@code createWorkbook} maps {@link PxlExcelEngine} to the POI implementation - {@code HSSF} to an
 * {@link HSSFWorkbook} (XLS), {@code XSSF} to an {@link XSSFWorkbook} (XLSX), {@code SXSSF} to an auto-flushing
 * {@link SXSSFWorkbook} (streaming XLSX) - and stamps the creator/application document properties.
 * <p>
 * {@code openWorkbook} is where the reader is chosen. It sniffs the source's leading magic bytes rather than
 * trusting the file extension, and reads with the streaming reader only when the import metadata asks for it
 * <em>and</em> the content really is OOXML; a file that turns out to be OLE2 (XLS) falls back to the non-streaming
 * reader, while a stream that does can no longer be rewound and therefore fails with a clear exception instead.
 * The format it settled on is recorded back into the metadata, and every POI failure is translated into
 * {@link InvalidFormatException}/{@link IOException} carrying a localized message.
 * <p>
 * The naming helpers exist because Excel constrains both: a sheet name is limited to 31 characters and forbids
 * certain characters, and neither a sheet name nor a defined name may repeat within a workbook. Both helpers
 * sanitize first and then append a {@code " (2)"}, {@code " (3)"} ... suffix until the name is free.
 */
public final class PxlWorkbookSupport {

    /**
     * Prevents instantiation.
     */
    private PxlWorkbookSupport() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Obtains the field holding the workbook name from the workbook class.
     * <p>
     * Scans the whole inheritance chain and returns the first field annotated with {@link PxlWorkbookName}.
     *
     * @param workbookClass the workbook class to scan
     * @return the {@link PxlWorkbookName}-annotated field, or {@code null} if none exists
     */
    public static Field getWorkbookNameField(final Class<?> workbookClass) {

//        return Arrays.stream(workbookClass.getDeclaredFields())
//                .filter(o -> Objects.nonNull(o.getAnnotation(PxlWorkbookName.class)))
//                .findFirst()
//                .orElse(null);
        return PxlReflectionSupport.getAllFields(workbookClass).stream()
                .filter(o -> Objects.nonNull(o.getAnnotation(PxlWorkbookName.class)))
                .findFirst()
                .orElse(null);
    }

    /**
     * Validates that the {@link PxlWorkbookName}-annotated field (if any) is declared as {@link String}.
     * No-ops when {@code workbookClass} is {@code null} or has no such field.
     *
     * @param workbookClass the workbook class to validate
     * @throws PxlDataException when the annotated field's type is not {@link String}
     */
    public static void validateWorkbookNameFieldType(final Class<?> workbookClass)
            throws PxlDataException {

        if (Objects.isNull(workbookClass)) {
            return;
        }

        final Field workbookNameField = getWorkbookNameField(workbookClass);
        if (Objects.nonNull(workbookNameField) && workbookNameField.getType() != String.class) {
            throw new PxlDataException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_WORKBOOK_NAME_STRING_ONLY,
                    workbookNameField.getName(), workbookNameField.getType().getSimpleName()));
        }
    }

    /**
     * Creates an empty workbook for the requested export engine and stamps its creator/application document properties.
     * <p>
     * {@code HSSF} produces an {@link HSSFWorkbook} (XLS), {@code XSSF} an {@link XSSFWorkbook} (XLSX), and {@code SXSSF} an
     * auto-flushing {@link SXSSFWorkbook} (streaming XLSX) with the given row-access window size and temp-file compression off.
     * <p>
     * Every {@link PxlExcelEngine} constant is an Excel writer, so no argument can name an output this is unable
     * to create - CSV, which has no POI workbook, is not an engine. There is consequently no unsupported-format
     * failure to report.
     *
     * @param exportExcelEngine              the target POI engine (HSSF, XSSF, or SXSSF)
     * @param exportSXSSFRowAccessWindowSize the in-memory row window size, used only for the SXSSF (streaming) engine
     * @return the created workbook
     */
    public static Workbook createWorkbook(final PxlExcelEngine exportExcelEngine,
                                          final int exportSXSSFRowAccessWindowSize) {

        Workbook workbook;

        switch (exportExcelEngine) {
            case HSSF: {
                workbook = new HSSFWorkbook();

                ((HSSFWorkbook) workbook).createInformationProperties();
                ((HSSFWorkbook) workbook).getSummaryInformation().setAuthor(PxlConstants.PXL_CREATOR);
                ((HSSFWorkbook) workbook).getSummaryInformation().setApplicationName(PxlConstants.PXL_APPLICATION);

                break;
            }

            case XSSF: {
                workbook = new XSSFWorkbook();

                final POIXMLProperties properties = ((XSSFWorkbook) workbook).getProperties();
                properties.getCoreProperties().setCreator(PxlConstants.PXL_CREATOR);
                properties.getExtendedProperties().getUnderlyingProperties().setApplication(PxlConstants.PXL_APPLICATION);

                break;
            }

            case SXSSF: {
                // turn on auto-flushing, no need to call SXSSFSheet::flushRows()
                workbook = new SXSSFWorkbook(exportSXSSFRowAccessWindowSize);

                ((SXSSFWorkbook) workbook).setCompressTempFiles(false/*compressTmpFiles*/);
                final POIXMLProperties properties = ((SXSSFWorkbook) workbook).getXSSFWorkbook().getProperties();
                properties.getCoreProperties().setCreator(PxlConstants.PXL_CREATOR);
                properties.getExtendedProperties().getUnderlyingProperties().setApplication(PxlConstants.PXL_APPLICATION);

                break;
            }

            default:
                // Unreachable: every PxlExcelEngine constant is handled above. Kept so that adding an engine
                // without extending this switch fails loudly instead of returning null.
                throw new AssertionError(exportExcelEngine);
        }
        //WorkbookFactory.create(xssf);

        return workbook;
    }

    /**
     * Opens an Excel workbook from a file, choosing the reader based on the file's magic bytes and the import metadata.
     * <p>
     * When {@code workbookMeta} is {@code null}, opens read-only via {@link WorkbookFactory}. Otherwise, when streaming
     * is requested and the file is OOXML (XLSX), opens with {@link StreamingReader} (falling back to non-streaming
     * {@link WorkbookFactory} if the content turns out to be an OLE2/XLS file); otherwise uses {@link WorkbookFactory}
     * with the import password. The resolved format (HSSF vs XSSF) is recorded back into {@code workbookMeta}.
     *
     * @param workbookFile the Excel file to open
     * @param workbookMeta the import metadata (password, streaming options); may be {@code null} for a plain read-only open
     * @return the opened workbook
     * @throws InvalidFormatException when the format is unsupported or the password is wrong
     * @throws IOException            when the file does not exist or cannot be read
     */
    public static Workbook openWorkbook(final File workbookFile,
                                        final PxlImportWorkbookMeta workbookMeta)
            throws IOException, InvalidFormatException {

        try {
            Workbook workbook = null;

            if (Objects.isNull(workbookMeta)) {
                workbook = WorkbookFactory.create(workbookFile, null, true);
            } else {
                final String importPassword = StringUtils.defaultIfEmpty(workbookMeta.getImportPassword(), null);

                // Excel Stream Reader only supports reading XLSX files.
                final boolean importUsingStreamReader = workbookMeta.isImportUsingStreamReader();

                // Determine the file type first from the leading magic bytes (up to 8 bytes).
                final FileMagic fileMagic = FileMagic.valueOf(workbookFile);

                if (importUsingStreamReader && fileMagic == FileMagic.OOXML) {
                    try {
                        final int importStreamReaderRowCacheSize = workbookMeta.getImportStreamReaderRowCacheSize();
                        final int importStreamReaderBufferSize = workbookMeta.getImportStreamReaderBufferSize();

                        workbook = StreamingReader.builder()
                                .rowCacheSize(importStreamReaderRowCacheSize)
                                .bufferSize(importStreamReaderBufferSize)
                                .password(importPassword)
                                .open(workbookFile);
                    } catch (OLE2NotOfficeXmlFileException e) {
                        PxlWorkbookUtils.closeWorkbook(workbook);
                        workbook = WorkbookFactory.create(workbookFile, importPassword, true);
                    } catch (ReadException e) {
                        PxlWorkbookUtils.closeWorkbook(workbook);
                        if (e.getCause() instanceof OLE2NotOfficeXmlFileException) {
                            workbook = WorkbookFactory.create(workbookFile, importPassword, true);
                        } else {
                            throw e;
                        }
                    }
                } else {
                    workbook = WorkbookFactory.create(workbookFile, importPassword, true);
                }

                if (workbook instanceof HSSFWorkbook) {
                    workbookMeta.setImportFileFormat(PxlFileFormat.XLS);
                } else {
                    workbookMeta.setImportFileFormat(PxlFileFormat.XLSX);
                }
            }

            return workbook;
        } catch (EncryptedDocumentException e) {
            throw new InvalidFormatException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_WORKBOOK_DECRYPT_FAILED), e);
        } catch (UnsupportedFileFormatException e) {
            // NOTE: the file may be DRM-protected.
            throw new InvalidFormatException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_WORKBOOK_UNSUPPORTED_FORMAT));
        } catch (FileNotFoundException e) {
            throw new FileNotFoundException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_WORKBOOK_FILE_NOT_FOUND, e.getMessage()));
        } catch (IOException | OpenException | ReadException e) {
            throw new IOException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_WORKBOOK_FILE_UNREADABLE, e.getMessage()));
        }
    }

    /**
     * Opens an Excel workbook from an input stream, choosing the reader based on the stream's magic bytes and the import metadata.
     * <p>
     * When {@code workbookMeta} is {@code null}, opens via {@link WorkbookFactory}. Otherwise, when streaming is requested and
     * the stream is OOXML (XLSX), opens with {@link StreamingReader} over a {@link CloseShieldInputStream} so the caller's
     * stream is not closed; otherwise uses {@link WorkbookFactory} with the import password. Because a stream already consumed
     * by the streaming reader cannot be rewound, an OLE2/XLS stream fails with {@link InvalidFormatException} rather than
     * falling back. The resolved format (HSSF vs XSSF) is recorded back into {@code workbookMeta}.
     *
     * @param inputStream  the stream to read the workbook from
     * @param workbookMeta the import metadata (password, streaming options); may be {@code null} for a plain open
     * @return the opened workbook
     * @throws InvalidFormatException when the format is unsupported, the password is wrong, or a consumed stream is not XLSX
     * @throws IOException            when the stream cannot be read
     */
    public static Workbook openWorkbook(final InputStream inputStream,
                                        final PxlImportWorkbookMeta workbookMeta)
            throws IOException, InvalidFormatException {

        try {
            final BufferedInputStream bufferedInputStream = new BufferedInputStream(inputStream);

            Workbook workbook = null;

            if (Objects.isNull(workbookMeta)) {
                // Shield the caller-owned original stream from being closed when WorkbookFactory closes the workbook.
                final CloseShieldInputStream csis = CloseShieldInputStream.wrap(bufferedInputStream);
                workbook = WorkbookFactory.create(csis);
            } else {
                final String importPassword = StringUtils.defaultIfEmpty(workbookMeta.getImportPassword(), null);

                // Excel Stream Reader only supports reading XLSX files.
                final boolean importUsingStreamReader = workbookMeta.isImportUsingStreamReader();

                // Determine the stream type first from the leading magic bytes (up to 8 bytes).
                final InputStream magicCheckedStream = FileMagic.prepareToCheckMagic(bufferedInputStream);
                final FileMagic fileMagic = FileMagic.valueOf(magicCheckedStream);

                if (importUsingStreamReader && fileMagic == FileMagic.OOXML) {
                    try {
                        final int importStreamReaderRowCacheSize = workbookMeta.getImportStreamReaderRowCacheSize();
                        final int importStreamReaderBufferSize = workbookMeta.getImportStreamReaderBufferSize();

                        // Shield the caller-owned original stream from being closed when StreamingReader closes the workbook.
                        final CloseShieldInputStream csis = CloseShieldInputStream.wrap(magicCheckedStream);

                        workbook = StreamingReader.builder()
                                .rowCacheSize(importStreamReaderRowCacheSize)
                                .bufferSize(importStreamReaderBufferSize)
                                .password(importPassword)
                                .open(csis);
                    } catch (OLE2NotOfficeXmlFileException e) {
                        PxlWorkbookUtils.closeWorkbook(workbook);
                        // An InputStream already consumed by the streaming reader cannot be rewound, so it cannot be re-read non-streaming.
                        // Fail with a clear exception instead of a broken fallback (re-reading a consumed stream).
                        // workbook = WorkbookFactory.create(magicCheckedStream, importPassword);
                        throw new InvalidFormatException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_WORKBOOK_STREAMING_UNREADABLE), e);
                    } catch (ReadException e) {
                        PxlWorkbookUtils.closeWorkbook(workbook);
                        if (e.getCause() instanceof OLE2NotOfficeXmlFileException) {
                            // workbook = WorkbookFactory.create(magicCheckedStream, importPassword);
                            throw new InvalidFormatException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_WORKBOOK_STREAMING_UNREADABLE), e);
                        } else {
                            throw e;
                        }
                    }
                } else {
                    // Shield the caller-owned original stream from being closed when WorkbookFactory closes the workbook.
                    final CloseShieldInputStream csis = CloseShieldInputStream.wrap(magicCheckedStream);

                    workbook = WorkbookFactory.create(csis, importPassword);
                }

                if (workbook instanceof HSSFWorkbook) {
                    workbookMeta.setImportFileFormat(PxlFileFormat.XLS);
                } else {
                    workbookMeta.setImportFileFormat(PxlFileFormat.XLSX);
                }
            }

            return workbook;
        } catch (EncryptedDocumentException e) {
            throw new InvalidFormatException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_WORKBOOK_DECRYPT_FAILED), e);
        } catch (UnsupportedFileFormatException e) {
            // NOTE: the file may be DRM-protected.
            throw new InvalidFormatException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_WORKBOOK_UNSUPPORTED_FORMAT));
        } catch (IOException | ReadException e) {
            throw new IOException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_WORKBOOK_FILE_UNREADABLE, e.getMessage()));
        }
    }

    /**
     * Returns the sheet names that collide with an earlier name in the given sequence, so an export can reject them
     * before writing anything.
     *
     * <p>Names are compared as the safe names they will become ({@link WorkbookUtil#createSafeSheetName(String)} -
     * invalid characters replaced, truncated to 31 chars), <strong>ignoring case</strong>: a workbook cannot hold two
     * sheets whose names differ only in case, which is also how a sheet name is matched on import. The comparison is
     * locale-independent. Sanitizing is idempotent, so a caller may pass names that are already safe.</p>
     *
     * <p>The returned names are the <em>original</em> ones, in encounter order, keeping the diagnostic message
     * pointing at what the caller actually declared.</p>
     *
     * @param desiredNames the sheet names about to be written, in order; may be {@code null}
     * @return the colliding names (every occurrence after the first), empty when all are unique
     */
    public static Set<String> findDuplicateSheetNames(final Collection<String> desiredNames) {

        final Set<String> safeNames = new TreeSet<>(String.CASE_INSENSITIVE_ORDER);
        return PxlCollectionUtils.emptyIfNull(desiredNames).stream()
                .filter(Objects::nonNull)
                .filter(desiredName -> !safeNames.add(WorkbookUtil.createSafeSheetName(desiredName)))
                .collect(Collectors.toCollection(LinkedHashSet::new));
    }

    /**
     * Creates a safe sheet name (<=31 chars, no invalid characters) that does not collide with any sheet name already present in the workbook.
     * On a collision, appends a " (2)", " (3)" ... suffix while staying within the 31-char limit. (POI's getSheet matches case-insensitively.)
     *
     * @param workbook    the target workbook
     * @param desiredName the desired sheet name
     * @return a safe sheet name unique within the workbook
     */
    public static String makeUniqueSafeSheetName(final Workbook workbook,
                                                 final String desiredName) {

        final String safeName = WorkbookUtil.createSafeSheetName(desiredName);
        if (Objects.isNull(workbook.getSheet(safeName))) {
            return safeName;
        }

        for (int suffixNumber = 2; ; suffixNumber++) {
            final String suffix = " (" + suffixNumber + ")";
            final int maxBaseLength = PxlConstants.MAX_SHEET_NAME_LENGTH - suffix.length();
            final String base = safeName.length() > maxBaseLength ? safeName.substring(0, maxBaseLength) : safeName;
            final String candidate = base + suffix;   // both base and suffix are already safe, so no re-sanitization is needed.
            if (Objects.isNull(workbook.getSheet(candidate))) {
                return candidate;
            }
        }
    }

    /**
     * Creates a unique name that does not collide with any defined name (Named Range) already present in the workbook.
     * On a collision, appends a "_2", "_3" ... suffix.
     *
     * @param workbook    the target workbook
     * @param desiredName the desired name
     * @return a name unique within the workbook
     */
    public static String makeUniqueDefinedName(final Workbook workbook,
                                               final String desiredName) {

        if (Objects.isNull(workbook.getName(desiredName))) {
            return desiredName;
        }

        for (int suffixNumber = 2; ; suffixNumber++) {
            final String candidate = desiredName + "_" + suffixNumber;
            if (Objects.isNull(workbook.getName(candidate))) {
                return candidate;
            }
        }
    }

    /**
     * Tests whether the string is a syntactically valid filesystem path.
     *
     * @param path the path string to test
     * @return {@code true} if the string parses as a path, {@code false} if it is {@code null} or invalid
     */
    @Deprecated
    private static boolean isValidPath(final String path) {

        try {
            Paths.get(path);
        } catch (InvalidPathException | NullPointerException ignored) {
            return false;
        }
        return true;
    }

    /**
     * Tests whether the file is an OLE2 container (as opposed to an OOXML file), by attempting to open it as a {@link POIFSFileSystem}.
     *
     * @param file the file to inspect
     * @return {@code true} if the file is an OLE2 container, {@code false} if it is an OOXML file
     * @throws IOException if the file cannot be read
     */
    @Deprecated
    private static boolean isEncrypted(final File file)
            throws IOException {

        try (final POIFSFileSystem fileSystem = new POIFSFileSystem(file)) {
            return true;
        } catch (OfficeXmlFileException e) {
            return false;
        }
    }

    /**
     * Tests whether the stream holds an OLE2 container (as opposed to an OOXML file), by attempting to open it as a {@link POIFSFileSystem}.
     *
     * @param is the stream to inspect
     * @return {@code true} if the stream is an OLE2 container, {@code false} if it is an OOXML file
     * @throws IOException if the stream cannot be read
     */
    @Deprecated
    private static boolean isEncrypted(final InputStream is)
            throws IOException {

        try (final POIFSFileSystem fileSystem = new POIFSFileSystem(is)) {
            return true;
        } catch (OfficeXmlFileException e) {
            return false;
        }
    }

    /**
     * Tests whether the file is an OOXML (XLSX) file, based on its magic bytes.
     *
     * @param file the file to inspect
     * @return {@code true} if the file is OOXML
     * @throws IOException if the file cannot be read
     */
    @Deprecated
    private static boolean isXLSX(final File file)
            throws IOException {

        return FileMagic.OOXML.equals(FileMagic.valueOf(file));
    }

    /**
     * Tests whether the stream holds an OOXML (XLSX) file, based on its magic bytes.
     *
     * @param is the stream to inspect
     * @return {@code true} if the stream is OOXML
     * @throws IOException if the stream cannot be read
     */
    @Deprecated
    private static boolean isXLSX(final InputStream is)
            throws IOException {

        return FileMagic.OOXML.equals(FileMagic.valueOf(is));
    }

}
