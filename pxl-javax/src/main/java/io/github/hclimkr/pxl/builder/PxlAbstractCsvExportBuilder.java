package io.github.hclimkr.pxl.builder;

import io.github.hclimkr.pxl.PxlConstants;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlException;
import io.github.hclimkr.pxl.exception.PxlIOException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlExportWorkbookMeta;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.option.PxlExportSheetOption;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.lang3.StringUtils;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * CSV-specific base for the export builders, filling in the three seams of {@link PxlAbstractExportBuilder}.
 *
 * <p>{@link #prepare()} renders the whole CSV into an in-memory buffer and {@link #writeTo(OutputStream)} only
 * copies it out. Rendering ahead of the destination being opened is what gives CSV the same failure semantics as
 * Excel: a codec, validation or limit failure leaves no file behind. The price is that the memory needed grows
 * with the output.</p>
 *
 * <p>The division of labour with the core is that the charset, the byte order mark and the buffer belong here,
 * while the CSV grammar belongs to {@code PxlCoreCsvExporter}. That is why the seam subclasses fill in receives a
 * {@link Writer} rather than a printer.</p>
 *
 * <p>Package-private: not part of the public API.</p>
 */
abstract class PxlAbstractCsvExportBuilder extends PxlAbstractExportBuilder {

    /**
     * Sheet names, in the order {@code sheet(...)} was called. CSV writes one file per sheet, so the terminals
     * accept exactly one; the check runs in {@link #prepare()} rather than in {@code sheet(...)} because it is the
     * terminal, not the builder, that cannot take more than one.
     */
    protected final List<String> sheetNames = new ArrayList<>();

    /**
     * Row classes, aligned with {@link #sheetNames}.
     */
    protected final List<Class<?>> rowClasses = new ArrayList<>();

    /**
     * The rendered output, held between {@code prepare()} and {@code writeTo(...)}.
     */
    private ByteArrayOutputStream buffer;

    /**
     * Validates the configuration and renders the whole CSV into memory.
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

        final ByteArrayOutputStream rendered = new ByteArrayOutputStream();

        try {
            final Writer writer = new OutputStreamWriter(rendered, charset);

            if (resolveCsvBom(workbookMeta) && writesBom(charset)) {
                writer.write('﻿');
            }

            writeRecords(writer, workbookMeta);

            // Push the encoder's remainder into the buffer; the core only flushed as far as this writer.
            writer.flush();
        } catch (IOException e) {
            throw new PxlIOException(e);
        }

        this.buffer = rendered;
    }

    /**
     * Copies the rendered output to the destination.
     *
     * @param outputStream the destination output stream
     * @throws PxlIOException if writing fails
     */
    @Override
    protected final void writeTo(final OutputStream outputStream)
            throws PxlIOException {

        try {
            // Not toByteArray(), which would copy the whole output once more.
            buffer.writeTo(outputStream);
        } catch (IOException e) {
            throw new PxlIOException(e);
        }
    }

    /**
     * Releases the rendered output so it is not held until the next run.
     */
    @Override
    protected final void cleanup() {

        this.buffer = null;
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
     * @param workbookMeta the resolved export metadata for the workbook
     * @param sheetName    the sheet name named in the failure message
     * @throws PxlArgumentException if the delimiter cannot be used
     */
    private void assertUsableDelimiter(final PxlExportWorkbookMeta workbookMeta,
                                       final String sheetName)
            throws PxlArgumentException {

        final char delimiter = resolveCsvDelimiter(workbookMeta);

        try {
            CSVFormat.EXCEL.builder().setQuote('"').setDelimiter(delimiter).build();
        } catch (IllegalArgumentException e) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.BUILDER_EXPORT_CSV_DELIMITER_INVALID, sheetName, String.valueOf(delimiter)));
        }
    }

    /**
     * Resolves the sheet-level CSV charset the way the ad-hoc sheet meta does — the wildcard sheet option first,
     * then the workbook value. The builder needs it before the core builds the sheet meta, since it is the builder
     * that encodes.
     *
     * @param workbookMeta the resolved export metadata for the workbook
     * @return the charset name in effect
     */
    private String resolveCsvCharsetName(final PxlExportWorkbookMeta workbookMeta) {

        return Optional.ofNullable(findWildcardSheetOption(workbookMeta))
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

        return Optional.ofNullable(findWildcardSheetOption(workbookMeta))
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

        return Optional.ofNullable(findWildcardSheetOption(workbookMeta))
                .map(PxlExportSheetOption::getExportCsvBom)
                .orElseGet(workbookMeta::isExportCsvBom);
    }

    /**
     * Returns the wildcard sheet option, the only sheet-level override reachable from the sheet form.
     *
     * @param workbookMeta the resolved export metadata for the workbook
     * @return the wildcard sheet option, or {@code null} if none is registered
     */
    private PxlExportSheetOption findWildcardSheetOption(final PxlExportWorkbookMeta workbookMeta) {

        return Optional.ofNullable(workbookMeta.getExportSheetOptions())
                .flatMap(options -> options.stream()
                        .filter(o -> StringUtils.equals(o.getFieldName(), PxlConstants.SHEET_FIELD_NAME_WILD_CARD))
                        .findFirst())
                .orElse(null);
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
