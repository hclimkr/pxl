package io.github.hclimkr.pxl.builder;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlException;
import io.github.hclimkr.pxl.exception.PxlIOException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlExportWorkbookMeta;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.internal.support.PxlOptionSupport;
import io.github.hclimkr.pxl.option.PxlExportSheetOption;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.io.output.DeferredFileOutputStream;
import org.apache.commons.lang3.StringUtils;

import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * CSV-specific base for the export builders, filling in the three seams of {@link PxlAbstractExportBuilder}.
 *
 * <p>{@link #prepare()} renders the whole CSV before {@link #writeTo(OutputStream)} copies it out. Rendering ahead
 * of the destination being opened is what gives CSV the same failure semantics as Excel: a codec, validation or
 * limit failure leaves no file behind.</p>
 *
 * <p>Where that rendering is held depends on how large it turns out to be. Up to
 * {@link PxlConstants#EXPORT_MEMORY_THRESHOLD_OF_CSV} it stays in memory, and beyond it the rest continues
 * into a temporary file, which is what {@link DeferredFileOutputStream} switches between. Holding all of it in
 * memory made the heap the ceiling - the buffer's growth copy puts the peak at two to three times the output - while
 * always using a file would charge every small export a temporary file it does not need. The spill trades heap for
 * disk without touching the failure semantics: the destination is still opened only once the output is complete.</p>
 *
 * <p>Because the sink may hold a temporary file, {@link #cleanup()} rather than {@code prepare()} owns its removal,
 * and the field is assigned before rendering starts so that a failure part-way through still has something to
 * release.</p>
 *
 * <p>The division of labour with the core is that the charset, the byte order mark and the sink belong here,
 * while the CSV grammar belongs to {@code PxlCoreCsvExporter}. That is why the seam subclasses fill in receives a
 * {@link Writer} rather than a printer.</p>
 *
 * <p>Package-private: not part of the public API.</p>
 */
abstract class PxlAbstractCsvExportBuilder extends PxlAbstractExportBuilder {

    /**
     * The byte order mark, U+FEFF. Written as a code point rather than as the character itself, which is
     * zero-width: a literal one is invisible in the source and does not survive a tool that reads or rewrites the
     * file in a non-Unicode charset.
     */
    private static final char BOM = (char) 0xFEFF;

    /**
     * Name prefix of the temporary file a large export spills into. Distinctive enough that a file left behind by a
     * killed JVM can be told apart from other temporary files.
     */
    private static final String TEMP_FILE_PREFIX = "pxl-csv-export-";

    /**
     * Name suffix of that temporary file. Deliberately not {@code .csv}: the file is an implementation detail, and
     * a half-written one must not look like a result.
     */
    private static final String TEMP_FILE_SUFFIX = ".tmp";

    /**
     * Sheet names, in the order {@code sheet(...)} was called. CSV writes one file per sheet, so the terminals
     * accept exactly one; the check runs in {@link #prepare()} rather than in {@code sheet(...)} because it is the
     * terminal, not the builder, that cannot take more than one.
     */
    protected final List<String> sheetNames = new ArrayList<>();

    /**
     * Row classes, aligned with {@code sheetNames}.
     */
    protected final List<Class<?>> rowClasses = new ArrayList<>();

    /**
     * The rendered output, held between {@code prepare()} and {@code writeTo(...)}. In memory while it is small
     * enough, in a temporary file once it is not.
     */
    private DeferredFileOutputStream deferredFileOutputStream;

    /**
     * Validates the configuration and renders the whole CSV, in memory up to
     * {@link PxlConstants#EXPORT_MEMORY_THRESHOLD_OF_CSV} and into a temporary file beyond it.
     *
     * @throws PxlArgumentException if no sheet or more than one sheet is configured, a password is requested, or the
     *                              charset/delimiter cannot be used
     * @throws PxlException         if the metadata cannot be resolved or a record cannot be written
     */
    @Override
    protected final void prepare()
            throws PxlException {

        PxlAssertSupport.notEmpty(sheetNames, "sheetName");

        if (PxlCollectionUtils.size(sheetNames) > 1) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.BUILDER_EXPORT_CSV_SINGLE_SHEET_ONLY));
        }

        final PxlExportWorkbookMeta workbookMeta = PxlExportWorkbookMeta.makeExportWorkbookMetaForCsv(null, option);

        final String sheetName = PxlCollectionUtils.get(sheetNames, 0);

        assertNoPassword(workbookMeta);

        final Charset charset = resolveCharset(workbookMeta, sheetName);
        assertUsableDelimiter(workbookMeta, sheetName);

        // Assigned before rendering: a failure part-way through may already have spilled to a file, and cleanup()
        // can only remove what it can reach.
        deferredFileOutputStream = DeferredFileOutputStream.builder()
                .setThreshold(PxlConstants.EXPORT_MEMORY_THRESHOLD_OF_CSV)
                .setPrefix(TEMP_FILE_PREFIX)
                .setSuffix(TEMP_FILE_SUFFIX)
                .get();

        try {
            final Writer writer = new OutputStreamWriter(deferredFileOutputStream, charset);

            if (resolveCsvBom(workbookMeta) && writesBom(charset)) {
                writer.write(BOM);
            }

            writeRecords(writer, workbookMeta);

            // Push the encoder's remainder into the sink; the core only flushed as far as this writer.
            writer.flush();

            // writeTo(...) refuses to read a sink that is still open, and a spilled one has to reach the disk first.
            deferredFileOutputStream.close();
        } catch (IOException e) {
            throw new PxlIOException(e);
        }
    }

    /**
     * Copies the rendered output to the destination, from memory or from the temporary file it spilled into.
     *
     * @param outputStream the destination output stream
     * @throws PxlIOException if writing fails
     */
    @Override
    protected final void writeTo(final OutputStream outputStream)
            throws PxlIOException {

        try {
            // Not getData(), which would copy the whole output once more - and which answers null once spilled.
            deferredFileOutputStream.writeTo(outputStream);
        } catch (IOException e) {
            throw new PxlIOException(e);
        }
    }

    /**
     * Releases the rendered output so neither the memory nor the temporary file outlives the run.
     *
     * <p>Runs on the failure path too, which is the only thing standing between a mid-render failure and a
     * temporary file left on disk.</p>
     */
    @Override
    protected final void cleanup() {

        final DeferredFileOutputStream sink = deferredFileOutputStream;
        deferredFileOutputStream = null;

        if (Objects.isNull(sink)) {
            return;
        }

        // Closing twice is harmless, and a sink abandoned mid-render was never closed at all.
        IOUtils.closeQuietly(sink);

        // null while the output stayed in memory; a temporary file only exists once the threshold was passed.
        FileUtils.deleteQuietly(sink.getFile());
    }

    /**
     * Writes this builder's records through the given writer. The only seam the concrete CSV builders fill in.
     *
     * @param writer       the destination the records are printed to
     * @param workbookMeta the resolved export metadata for the workbook
     * @throws PxlException if a record cannot be written
     */
    protected abstract void writeRecords(Writer writer, PxlExportWorkbookMeta workbookMeta)
            throws PxlException;

    /**
     * Rejects a password, which CSV cannot honor.
     *
     * <p>Every other Excel-only setting is ignored on this path, but silently writing plaintext when encryption was
     * asked for is a leak rather than a missing feature.</p>
     *
     * @param workbookMeta the resolved export metadata for the workbook
     * @throws PxlArgumentException if a password is configured
     */
    private static void assertNoPassword(final PxlExportWorkbookMeta workbookMeta)
            throws PxlArgumentException {

        if (StringUtils.isNotEmpty(workbookMeta.getExportPassword())) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.BUILDER_EXPORT_CSV_PASSWORD_UNSUPPORTED));
        }
    }

    /**
     * Resolves the charset the output is encoded with, failing before the destination is opened when the name is
     * not one this JVM supports.
     *
     * @param workbookMeta the resolved export metadata for the workbook
     * @param sheetName    the sheet name named in the failure message
     * @return the resolved charset
     * @throws PxlArgumentException if the charset name is not supported
     */
    private Charset resolveCharset(final PxlExportWorkbookMeta workbookMeta,
                                   final String sheetName)
            throws PxlArgumentException {

        final String charsetName = resolveCsvCharsetName(workbookMeta);

        try {
            return Charset.forName(charsetName);
        } catch (RuntimeException e) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.BUILDER_EXPORT_CSV_CHARSET_INVALID, sheetName, charsetName));
        }
    }

    /**
     * Rejects a field delimiter Commons-CSV cannot build a format with, before the destination is opened.
     *
     * <p>The check is run against {@code PxlConstants.DEFAULT_EXPORT_CSV_FORMAT}, the same dialect the core builds
     * its {@code CSVPrinter} from. Commons-CSV judges a delimiter by whether it collides with the quote, escape or
     * comment character, so a check that reassembled the dialect for itself would go on validating the old one
     * once that constant changed.</p>
     *
     * @param workbookMeta the resolved export metadata for the workbook
     * @param sheetName    the sheet name named in the failure message
     * @throws PxlArgumentException if the delimiter cannot be used
     */
    private void assertUsableDelimiter(final PxlExportWorkbookMeta workbookMeta,
                                       final String sheetName)
            throws PxlArgumentException {

        final char delimiter = resolveCsvDelimiter(workbookMeta);

        try {
            PxlConstants.DEFAULT_EXPORT_CSV_FORMAT
                    .builder()
                    .setDelimiter(delimiter)
                    .build();
        } catch (IllegalArgumentException e) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.BUILDER_EXPORT_CSV_DELIMITER_INVALID, sheetName, String.valueOf(delimiter)));
        }
    }

    /**
     * Resolves the sheet-level CSV charset the way the ad-hoc sheet meta does - the wildcard sheet option first,
     * then the workbook value. The builder needs it before the core builds the sheet meta, since it is the builder
     * that encodes.
     *
     * @param workbookMeta the resolved export metadata for the workbook
     * @return the charset name in effect
     */
    private String resolveCsvCharsetName(final PxlExportWorkbookMeta workbookMeta) {

        return Optional.ofNullable(PxlOptionSupport.findExportWildcardSheetOption(workbookMeta.getExportSheetOptions()))
                .map(PxlExportSheetOption::getExportCsvCharset)
                .filter(StringUtils::isNotBlank)
                .orElseGet(workbookMeta::getExportCsvCharset);
    }

    /**
     * Resolves the sheet-level CSV field delimiter the same way {@link #resolveCsvCharsetName} resolves the charset.
     *
     * @param workbookMeta the resolved export metadata for the workbook
     * @return the field delimiter in effect
     */
    private char resolveCsvDelimiter(final PxlExportWorkbookMeta workbookMeta) {

        return Optional.ofNullable(PxlOptionSupport.findExportWildcardSheetOption(workbookMeta.getExportSheetOptions()))
                .map(PxlExportSheetOption::getExportCsvDelimiter)
                .filter(delimiter -> delimiter != PxlConstants.UNSPECIFIED_EXPORT_CSV_DELIMITER)
                .orElseGet(workbookMeta::getExportCsvDelimiter);
    }

    /**
     * Resolves whether a byte order mark precedes the output, the same way {@link #resolveCsvCharsetName} resolves
     * the charset. A mark belongs to the file rather than the schema, so the sheet level settles it.
     *
     * @param workbookMeta the resolved export metadata for the workbook
     * @return {@code true} if a mark was asked for
     */
    private boolean resolveCsvBom(final PxlExportWorkbookMeta workbookMeta) {

        return Optional.ofNullable(PxlOptionSupport.findExportWildcardSheetOption(workbookMeta.getExportSheetOptions()))
                .map(PxlExportSheetOption::getExportCsvBom)
                .orElseGet(workbookMeta::isExportCsvBom);
    }

    /**
     * Answers whether a byte order mark may be written for the given charset.
     *
     * <p>Only UTF-8, UTF-16LE and UTF-16BE qualify. The endian-detecting UTF-16 is excluded because its encoder
     * writes a mark of its own, and a non-Unicode charset is excluded because it cannot encode U+FEFF and would
     * replace it with {@code '?'}, corrupting the first field.</p>
     *
     * <p>The test is on the {@link Charset} rather than on its name: a user may well have written {@code "utf8"}
     * or another alias, and comparing names would drop the mark without saying so.</p>
     *
     * @param charset the charset the output is encoded with
     * @return {@code true} if a mark may be written
     */
    private static boolean writesBom(final Charset charset) {

        return Objects.equals(charset, StandardCharsets.UTF_8)
                || Objects.equals(charset, StandardCharsets.UTF_16LE)
                || Objects.equals(charset, StandardCharsets.UTF_16BE);
    }

}
