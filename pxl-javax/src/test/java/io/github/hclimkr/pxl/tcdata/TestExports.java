package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.builder.PxlCsvExportBuilder;
import io.github.hclimkr.pxl.builder.PxlExcelExportBuilder;
import io.github.hclimkr.pxl.builder.PxlSampleCsvExportBuilder;
import io.github.hclimkr.pxl.builder.PxlSampleExcelExportBuilder;
import io.github.hclimkr.pxl.exception.PxlException;
import io.github.hclimkr.pxl.util.PxlWorkbookUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.TestInfo;

import java.io.*;
import java.nio.file.Files;

/**
 * Runs a configured export builder against one {@link ExportDest} and hands the result back in a shape every
 * destination can be asserted on, so a single assertion serves the whole sweep.
 * <p>
 * There are two shapes, and the difference between them is the point of the {@code WORKBOOK} destination:
 * <ul>
 *     <li>{@link #emit} always returns bytes. For {@code WORKBOOK} that means writing the returned workbook out
 *     with {@link PxlWorkbookUtils#writeToStream(Workbook, OutputStream, String)} - the documented way to have the
 *     export password applied to a workbook the terminal handed over unencrypted.</li>
 *     <li>{@link #workbookOf} returns a POI workbook. For {@code FILE} and {@code STREAM} that is necessarily a
 *     workbook reopened from the written bytes; for {@code WORKBOOK} it is the live workbook the terminal built,
 *     because that object is what the terminal exists to produce. The caller closes it.</li>
 * </ul>
 * File artifacts are named per destination ({@code <method>_<DEST><ext>}), so the runs of one parameterized test
 * do not overwrite each other and each stays inspectable under {@code target/test-outputs/}.
 *
 * <p><strong>An encrypted export must use an overload that takes the password.</strong> The file and stream
 * terminals apply the builder's own {@code exportPassword} themselves, but {@code toWorkbook()} does not - it hands
 * the workbook over unencrypted by contract - so a {@code WORKBOOK} run given no password writes plaintext. That
 * failure is quiet in the worst way: plaintext opens under any password at all, so a test asserting that the wrong
 * password is rejected would simply stop asserting anything.</p>
 */
public final class TestExports {

    /**
     * File extension written by the XSSF/SXSSF engines.
     */
    public static final String XLSX = ".xlsx";

    /**
     * File extension written by the HSSF engine.
     */
    public static final String XLS = ".xls";

    /**
     * File extension written by the CSV export.
     */
    public static final String CSV = ".csv";

    private TestExports() {
    }

    /**
     * The {@code toFile(File)} terminal of some export builder.
     */
    @FunctionalInterface
    private interface FileTerminal {

        void toFile(File outputFile)
                throws PxlException;

    }

    /**
     * The {@code toStream(OutputStream)} terminal of some export builder.
     */
    @FunctionalInterface
    private interface StreamTerminal {

        void toStream(OutputStream outputStream)
                throws PxlException;

    }

    /**
     * The {@code toWorkbook()} terminal of an Excel export builder, or {@code null} for a CSV one.
     */
    @FunctionalInterface
    private interface WorkbookTerminal {

        Workbook toWorkbook()
                throws PxlException;

    }

    /**
     * The artifact a {@code FILE} run of the given test writes, named so the destinations of one parameterized
     * test never collide.
     *
     * @param testInfo  the current test (its method name is the artifact base name)
     * @param dest      the destination this run targets
     * @param extension the file extension, including the dot
     * @return the file handle under {@code target/test-outputs/}
     */
    public static File exportFile(final TestInfo testInfo, final ExportDest dest, final String extension) {

        return TestPaths.exportFile(testInfo, "_" + dest + extension);
    }

    // ------------------------------------------------------------------
    // Excel - data export
    // ------------------------------------------------------------------

    /**
     * Runs the builder against the destination and returns the exported {@code .xlsx} bytes.
     *
     * <p>For an encrypted export use the password overload instead - see the class javadoc.</p>
     *
     * @param builder  the configured builder
     * @param dest     the destination to drive it to
     * @param testInfo the current test, used to name the file artifact
     * @return the exported bytes
     * @throws PxlException if the export fails
     * @throws IOException  if the written artifact cannot be read back
     */
    public static byte[] emit(final PxlExcelExportBuilder builder,
                              final ExportDest dest,
                              final TestInfo testInfo)
            throws PxlException, IOException {

        return emit(builder, dest, testInfo, XLSX, null);
    }

    /**
     * Runs the builder against the destination and returns the exported bytes.
     *
     * <p>For an encrypted export use the password overload instead - see the class javadoc.</p>
     *
     * @param builder   the configured builder
     * @param dest      the destination to drive it to
     * @param testInfo  the current test, used to name the file artifact
     * @param extension the artifact extension, matching the engine the builder writes with
     * @return the exported bytes
     * @throws PxlException if the export fails
     * @throws IOException  if the written artifact cannot be read back
     */
    public static byte[] emit(final PxlExcelExportBuilder builder,
                              final ExportDest dest,
                              final TestInfo testInfo,
                              final String extension)
            throws PxlException, IOException {

        return emit(builder, dest, testInfo, extension, null);
    }

    /**
     * Runs the builder against the destination and returns the exported bytes, encrypting the {@code WORKBOOK}
     * result with the given password - the file and stream terminals apply the builder's own password themselves.
     *
     * @param builder   the configured builder
     * @param dest      the destination to drive it to
     * @param testInfo  the current test, used to name the file artifact
     * @param extension the artifact extension, matching the engine the builder writes with
     * @param password  the export password, or {@code null} when the export is not encrypted
     * @return the exported bytes
     * @throws PxlException if the export fails
     * @throws IOException  if the written artifact cannot be read back
     */
    public static byte[] emit(final PxlExcelExportBuilder builder,
                              final ExportDest dest,
                              final TestInfo testInfo,
                              final String extension,
                              final String password)
            throws PxlException, IOException {

        return emit(dest, exportFile(testInfo, dest, extension), password,
                builder::toFile, builder::toStream, builder::toWorkbook);
    }

    /**
     * Runs the builder against the destination and returns the result as a POI workbook (see the class javadoc
     * for what the {@code WORKBOOK} destination returns). The caller closes it.
     *
     * @param builder  the configured builder
     * @param dest     the destination to drive it to
     * @param testInfo the current test, used to name the file artifact
     * @return the exported workbook
     * @throws PxlException if the export fails or the result cannot be reopened
     * @throws IOException  if the written artifact cannot be read back
     */
    public static Workbook workbookOf(final PxlExcelExportBuilder builder,
                                      final ExportDest dest,
                                      final TestInfo testInfo)
            throws PxlException, IOException {

        return workbookOf(builder, dest, testInfo, XLSX, null);
    }

    /**
     * Runs the builder against the destination and returns the result as a POI workbook (see the class javadoc
     * for what the {@code WORKBOOK} destination returns). The caller closes it.
     *
     * @param builder   the configured builder
     * @param dest      the destination to drive it to
     * @param testInfo  the current test, used to name the file artifact
     * @param extension the artifact extension, matching the engine the builder writes with
     * @return the exported workbook
     * @throws PxlException if the export fails or the result cannot be reopened
     * @throws IOException  if the written artifact cannot be read back
     */
    public static Workbook workbookOf(final PxlExcelExportBuilder builder,
                                      final ExportDest dest,
                                      final TestInfo testInfo,
                                      final String extension)
            throws PxlException, IOException {

        return workbookOf(builder, dest, testInfo, extension, null);
    }

    /**
     * Runs the builder against the destination and returns the result as a POI workbook (see the class javadoc
     * for what the {@code WORKBOOK} destination returns). The caller closes it.
     *
     * @param builder   the configured builder
     * @param dest      the destination to drive it to
     * @param testInfo  the current test, used to name the file artifact
     * @param extension the artifact extension, matching the engine the builder writes with
     * @param password  the export password, or {@code null} when the export is not encrypted
     * @return the exported workbook
     * @throws PxlException if the export fails or the result cannot be reopened
     * @throws IOException  if the written artifact cannot be read back
     */
    public static Workbook workbookOf(final PxlExcelExportBuilder builder,
                                      final ExportDest dest,
                                      final TestInfo testInfo,
                                      final String extension,
                                      final String password)
            throws PxlException, IOException {

        if (ExportDest.WORKBOOK == dest) {
            return builder.toWorkbook();
        }

        return reopen(emit(builder, dest, testInfo, extension, password), password);
    }

    // ------------------------------------------------------------------
    // Excel - sample (template) export
    // ------------------------------------------------------------------

    /**
     * Runs the sample builder against the destination and returns the exported {@code .xlsx} bytes.
     *
     * <p>For an encrypted export use the password overload instead - see the class javadoc.</p>
     *
     * @param builder  the configured builder
     * @param dest     the destination to drive it to
     * @param testInfo the current test, used to name the file artifact
     * @return the exported bytes
     * @throws PxlException if the export fails
     * @throws IOException  if the written artifact cannot be read back
     */
    public static byte[] emit(final PxlSampleExcelExportBuilder builder,
                              final ExportDest dest,
                              final TestInfo testInfo)
            throws PxlException, IOException {

        return emit(builder, dest, testInfo, XLSX, null);
    }

    /**
     * Runs the sample builder against the destination and returns the exported bytes.
     *
     * <p>For an encrypted export use the password overload instead - see the class javadoc.</p>
     *
     * @param builder   the configured builder
     * @param dest      the destination to drive it to
     * @param testInfo  the current test, used to name the file artifact
     * @param extension the artifact extension, matching the engine the builder writes with
     * @return the exported bytes
     * @throws PxlException if the export fails
     * @throws IOException  if the written artifact cannot be read back
     */
    public static byte[] emit(final PxlSampleExcelExportBuilder builder,
                              final ExportDest dest,
                              final TestInfo testInfo,
                              final String extension)
            throws PxlException, IOException {

        return emit(builder, dest, testInfo, extension, null);
    }

    /**
     * Runs the sample builder against the destination and returns the exported bytes, encrypting the
     * {@code WORKBOOK} result with the given password.
     *
     * @param builder   the configured builder
     * @param dest      the destination to drive it to
     * @param testInfo  the current test, used to name the file artifact
     * @param extension the artifact extension, matching the engine the builder writes with
     * @param password  the export password, or {@code null} when the export is not encrypted
     * @return the exported bytes
     * @throws PxlException if the export fails
     * @throws IOException  if the written artifact cannot be read back
     */
    public static byte[] emit(final PxlSampleExcelExportBuilder builder,
                              final ExportDest dest,
                              final TestInfo testInfo,
                              final String extension,
                              final String password)
            throws PxlException, IOException {

        return emit(dest, exportFile(testInfo, dest, extension), password,
                builder::toFile, builder::toStream, builder::toWorkbook);
    }

    /**
     * Runs the sample builder against the destination and returns the result as a POI workbook (see the class
     * javadoc for what the {@code WORKBOOK} destination returns). The caller closes it.
     *
     * @param builder  the configured builder
     * @param dest     the destination to drive it to
     * @param testInfo the current test, used to name the file artifact
     * @return the exported workbook
     * @throws PxlException if the export fails or the result cannot be reopened
     * @throws IOException  if the written artifact cannot be read back
     */
    public static Workbook workbookOf(final PxlSampleExcelExportBuilder builder,
                                      final ExportDest dest,
                                      final TestInfo testInfo)
            throws PxlException, IOException {

        return workbookOf(builder, dest, testInfo, XLSX, null);
    }

    /**
     * Runs the sample builder against the destination and returns the result as a POI workbook (see the class
     * javadoc for what the {@code WORKBOOK} destination returns). The caller closes it.
     *
     * @param builder   the configured builder
     * @param dest      the destination to drive it to
     * @param testInfo  the current test, used to name the file artifact
     * @param extension the artifact extension, matching the engine the builder writes with
     * @return the exported workbook
     * @throws PxlException if the export fails or the result cannot be reopened
     * @throws IOException  if the written artifact cannot be read back
     */
    public static Workbook workbookOf(final PxlSampleExcelExportBuilder builder,
                                      final ExportDest dest,
                                      final TestInfo testInfo,
                                      final String extension)
            throws PxlException, IOException {

        return workbookOf(builder, dest, testInfo, extension, null);
    }

    /**
     * Runs the sample builder against the destination and returns the result as a POI workbook (see the class
     * javadoc for what the {@code WORKBOOK} destination returns). The caller closes it.
     *
     * @param builder   the configured builder
     * @param dest      the destination to drive it to
     * @param testInfo  the current test, used to name the file artifact
     * @param extension the artifact extension, matching the engine the builder writes with
     * @param password  the export password, or {@code null} when the export is not encrypted
     * @return the exported workbook
     * @throws PxlException if the export fails or the result cannot be reopened
     * @throws IOException  if the written artifact cannot be read back
     */
    public static Workbook workbookOf(final PxlSampleExcelExportBuilder builder,
                                      final ExportDest dest,
                                      final TestInfo testInfo,
                                      final String extension,
                                      final String password)
            throws PxlException, IOException {

        if (ExportDest.WORKBOOK == dest) {
            return builder.toWorkbook();
        }

        return reopen(emit(builder, dest, testInfo, extension, password), password);
    }

    // ------------------------------------------------------------------
    // CSV
    // ------------------------------------------------------------------

    /**
     * Runs the CSV builder against the destination and returns the exported bytes. CSV has no workbook terminal,
     * so the sweep has to be narrowed to {@code FILE} and {@code STREAM}.
     *
     * @param builder  the configured builder
     * @param dest     the destination to drive it to
     * @param testInfo the current test, used to name the file artifact
     * @return the exported bytes
     * @throws PxlException             if the export fails
     * @throws IOException              if the written artifact cannot be read back
     * @throws IllegalArgumentException if {@code dest} is {@code WORKBOOK}
     */
    public static byte[] emit(final PxlCsvExportBuilder builder,
                              final ExportDest dest,
                              final TestInfo testInfo)
            throws PxlException, IOException {

        return emit(dest, exportFile(testInfo, dest, CSV), null,
                builder::toFile, builder::toStream, null);
    }

    /**
     * Runs the sample CSV builder against the destination and returns the exported bytes. CSV has no workbook
     * terminal, so the sweep has to be narrowed to {@code FILE} and {@code STREAM}.
     *
     * @param builder  the configured builder
     * @param dest     the destination to drive it to
     * @param testInfo the current test, used to name the file artifact
     * @return the exported bytes
     * @throws PxlException             if the export fails
     * @throws IOException              if the written artifact cannot be read back
     * @throws IllegalArgumentException if {@code dest} is {@code WORKBOOK}
     */
    public static byte[] emit(final PxlSampleCsvExportBuilder builder,
                              final ExportDest dest,
                              final TestInfo testInfo)
            throws PxlException, IOException {

        return emit(dest, exportFile(testInfo, dest, CSV), null,
                builder::toFile, builder::toStream, null);
    }

    // ------------------------------------------------------------------
    // shared
    // ------------------------------------------------------------------

    /**
     * Drives one of the three terminals and returns what it produced as bytes.
     */
    private static byte[] emit(final ExportDest dest,
                               final File outputFile,
                               final String password,
                               final FileTerminal fileTerminal,
                               final StreamTerminal streamTerminal,
                               final WorkbookTerminal workbookTerminal)
            throws PxlException, IOException {

        switch (dest) {
            case FILE: {
                fileTerminal.toFile(outputFile);

                return Files.readAllBytes(outputFile.toPath());
            }
            case STREAM: {
                final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
                streamTerminal.toStream(outputStream);

                return outputStream.toByteArray();
            }
            default: {
                if (workbookTerminal == null) {
                    throw new IllegalArgumentException("this builder has no toWorkbook() terminal: " + dest);
                }

                // toWorkbook() hands the workbook over unencrypted, so the password is applied on the way out.
                final Workbook workbook = workbookTerminal.toWorkbook();
                try {
                    final ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
                    PxlWorkbookUtils.writeToStream(workbook, outputStream, password);

                    return outputStream.toByteArray();
                } finally {
                    PxlWorkbookUtils.closeWorkbook(workbook);
                }
            }
        }
    }

    /**
     * Reopens exported bytes as a POI workbook, decrypting where a password was used.
     */
    private static Workbook reopen(final byte[] bytes, final String password)
            throws PxlException {

        return PxlWorkbookUtils.openWorkbook(new ByteArrayInputStream(bytes), password);
    }

}
