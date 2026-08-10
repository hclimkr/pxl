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

import javax.validation.Validator;
import java.io.Writer;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * Builder that exports data to a CSV file/stream. Created via {@link Pxl#exportCsv()}.
 *
 * <p>CSV is one file per sheet, so this builder has the sheet form only - there is no workbook form to call.
 * {@code sheet(...)} accumulates just as the Excel builder's does, but the terminals write a single sheet, so
 * configuring more than one fails there.</p>
 *
 * <p>Settings that only describe how a cell looks or what a workbook contains - stylers, column widths, row
 * heights, freeze panes, auto-filters, dropdowns, the engine and its streaming window - are ignored, since CSV has
 * no way to carry them. Two of them are worth stating exactly, because the value still reaches the file:</p>
 * <ul>
 *   <li>{@code exportStringAsFormula}: nothing is evaluated, and the text is written as it stands, leading
 *       {@code =} and all.</li>
 *   <li>{@code exportStringAsPicture}: no picture is embedded, so what lands in the field is the image
 *       location the value held.</li>
 * </ul>
 *
 * <p>{@code exportPassword} is a different case again: CSV cannot be encrypted, and writing plaintext instead
 * would be a leak, so it is rejected rather than ignored.</p>
 *
 * <p>The terminal methods and resource handling are provided by {@link PxlAbstractCsvExportBuilder} and its
 * format-neutral base {@link PxlAbstractExportBuilder}.</p>
 *
 * <p>Example: {@code pxl.exportCsv().sheet(User.class, users, "Users").override(opt).toFile(file);}</p>
 */
public final class PxlCsvExportBuilder extends PxlAbstractCsvExportBuilder {

    private final Validator validator;

    private final List<Collection<?>> sheetObjects = new ArrayList<>();

    /**
     * Creates a CSV export builder with the given validator.
     *
     * @param validator the bean-validation validator, or {@code null} when bean validation is disabled
     */
    public PxlCsvExportBuilder(final Validator validator) {

        this.validator = validator;
    }

    /**
     * Sets the sheet to write.
     *
     * <p>Accumulates like the Excel builder's counterpart, but a CSV terminal writes a single sheet, so calling
     * this more than once makes {@code toFile(...)}/{@code toStream(...)} fail.</p>
     *
     * @param rowClass  the row class
     * @param rows      the row objects for this sheet
     * @param sheetName the sheet name; must not be blank
     * @param <T>       the row type
     * @return this builder
     * @throws PxlNullPointerException if {@code rowClass}, {@code rows}, or {@code sheetName} is {@code null}
     * @throws PxlArgumentException    if {@code sheetName} is blank
     */
    public <T> PxlCsvExportBuilder sheet(final Class<T> rowClass,
                                         final Collection<T> rows,
                                         final String sheetName)
            throws PxlNullPointerException, PxlArgumentException {

        PxlAssertSupport.notNull(rowClass, "rowClass");
        PxlAssertSupport.notNull(rows, "rows");
        PxlAssertSupport.notBlank(sheetName, "sheetName");

        this.sheetNames.add(sheetName);
        this.sheetObjects.add(rows);
        this.rowClasses.add(rowClass);
        return this;
    }

    /**
     * Overrides annotation-declared values with the given export option. (Optional)
     *
     * @param option the export option, or {@code null}
     * @return this builder
     */
    public PxlCsvExportBuilder override(@Nullable final PxlExportWorkbookOption option) {

        this.option = option;
        return this;
    }

    /**
     * Writes the configured sheet's header and data records.
     *
     * @param writer       the destination the records are printed to
     * @param workbookMeta the resolved export metadata for the workbook
     * @throws PxlException if a record cannot be written
     */
    @Override
    protected void writeRecords(final Writer writer,
                                final PxlExportWorkbookMeta workbookMeta)
            throws PxlException {

        PxlCoreCsvExporter.writeCsv(PxlCollectionUtils.get(sheetNames, 0),
                PxlCollectionUtils.get(sheetObjects, 0),
                PxlCollectionUtils.get(rowClasses, 0),
                workbookMeta,
                validator,
                writer);
    }

}
