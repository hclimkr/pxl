package io.github.hclimkr.pxl.builder;

import io.github.hclimkr.pxl.Pxl;
import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.core.PxlCoreCsvExporter;
import io.github.hclimkr.pxl.internal.meta.PxlExportWorkbookMeta;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;

import java.io.Writer;

/**
 * Builder that generates a CSV sample template from a class: a header record plus a single record filled from each
 * column's {@code exportSample} value. Created via {@link Pxl#exportSampleCsv()}.
 *
 * <p>Unlike {@link PxlCsvExportBuilder} this carries no validator: a sample has no data objects to validate, its
 * values being the declared sample strings, so {@code exportDataValidation} has nothing to act on.</p>
 *
 * <p>CSV is one file per sheet, so this builder has the sheet form only, and the terminals write a single sheet.</p>
 *
 * <p>Example: {@code pxl.exportSampleCsv().sheet(User.class, "Users").toFile(file);}</p>
 */
public final class PxlSampleCsvExportBuilder extends PxlAbstractCsvExportBuilder {

    /**
     * Creates a CSV sample export builder.
     */
    public PxlSampleCsvExportBuilder() {

        // Nothing to initialize: a sample carries no data objects and therefore no validator.
    }

    /**
     * Sets the sheet to write a sample for.
     *
     * <p>Accumulates like the Excel builder's counterpart, but a CSV terminal writes a single sheet, so calling
     * this more than once makes {@code toFile(...)}/{@code toStream(...)} fail.</p>
     *
     * @param rowClass  the row class
     * @param sheetName the sheet name; must not be blank
     * @return this builder
     * @throws PxlNullPointerException if {@code rowClass} or {@code sheetName} is {@code null}
     * @throws PxlArgumentException    if {@code sheetName} is blank
     */
    public PxlSampleCsvExportBuilder sheet(final Class<?> rowClass,
                                           final String sheetName)
            throws PxlNullPointerException, PxlArgumentException {

        PxlAssertSupport.notNull(rowClass, "rowClass");
        PxlAssertSupport.notBlank(sheetName, "sheetName");

        this.sheetNames.add(sheetName);
        this.rowClasses.add(rowClass);
        return this;
    }

    /**
     * Overrides annotation-declared values with the given export option. (Optional)
     *
     * @param option the export option, or {@code null}
     * @return this builder
     */
    public PxlSampleCsvExportBuilder override(@Nullable final PxlExportWorkbookOption option) {

        this.option = option;
        return this;
    }

    /**
     * Writes the configured sheet's header record and its single sample record.
     *
     * @param writer       the destination the records are printed to
     * @param workbookMeta the resolved export metadata for the workbook
     * @throws PxlException if a record cannot be written
     */
    @Override
    protected void writeRecords(final Writer writer,
                                final PxlExportWorkbookMeta workbookMeta)
            throws PxlException {

        PxlCoreCsvExporter.writeSampleCsv(PxlCollectionUtils.get(sheetNames, 0),
                PxlCollectionUtils.get(rowClasses, 0),
                workbookMeta,
                writer);
    }

}
