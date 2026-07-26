package io.github.hclimkr.pxl.builder;

import io.github.hclimkr.pxl.exception.PxlArgumentException;
import io.github.hclimkr.pxl.exception.PxlException;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.exception.PxlSystemException;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.core.PxlExcelExporter;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlExportWorkbookMeta;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.option.PxlExportWorkbookOption;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import io.github.hclimkr.pxl.util.PxlWorkbookUtils;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * Builder that creates an Excel sample workbook/file/stream from a class: a header row plus a single sample data
 * row filled from each column's {@code exportSample} value (not an empty form). Created via {@link io.github.hclimkr.pxl.Pxl#exportSampleExcel()}.
 *
 * <p>There are two usage forms.</p>
 * <ul>
 *   <li>{@code @PxlWorkbook} class form: {@link #workbook(Class)}</li>
 *   <li>Sheet form (multiple calls for multiple sheets): {@link #sheet(String, Class)}</li>
 * </ul>
 *
 * <p>The two forms are mutually exclusive — specifying both {@code workbook(...)} and {@code sheet(...)} throws an exception.</p>
 *
 * <p>The terminal methods (returning a workbook / file / stream) and resource handling are provided by {@link PxlAbstractExportBuilder}.</p>
 *
 * <p>Example: {@code pxl.exportSampleExcel().sheet("Users", User.class).sheet("Orders", Order.class).toFile(file);}</p>
 */
public final class PxlSampleExcelExportBuilder extends PxlAbstractExportBuilder {

    private Class<?> workbookClass;

    private final List<String> sheetNames = new ArrayList<>();
    private final List<Class<?>> rowClasses = new ArrayList<>();

    /**
     * Creates a sample Excel export builder.
     */
    public PxlSampleExcelExportBuilder() {
    }

    /**
     * Creates a sample from the {@code @PxlWorkbook} class. (Mutually exclusive with the sheet form)
     *
     * @param workbookClass the {@code @PxlWorkbook}-annotated class
     * @return this builder
     * @throws PxlNullPointerException if {@code workbookClass} is {@code null}
     */
    public PxlSampleExcelExportBuilder workbook(final Class<?> workbookClass)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(workbookClass, "workbookClass");

        this.workbookClass = workbookClass;
        return this;
    }

    /**
     * Adds a sheet sample. Calling multiple times produces multiple sheets. (Mutually exclusive with the workbook class form)
     *
     * @param sheetName the sheet name; must not be blank
     * @param rowClass  the row class
     * @return this builder
     * @throws PxlNullPointerException if {@code sheetName} or {@code rowClass} is {@code null}
     * @throws PxlArgumentException    if {@code sheetName} is blank
     */
    public PxlSampleExcelExportBuilder sheet(final String sheetName,
                                             final Class<?> rowClass)
            throws PxlNullPointerException, PxlArgumentException {

        PxlAssertSupport.notBlank(sheetName, "sheetName");
        PxlAssertSupport.notNull(rowClass, "rowClass");

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
    public PxlSampleExcelExportBuilder override(@Nullable final PxlExportWorkbookOption option) {

        this.option = option;
        return this;
    }

    /**
     * Builds the Excel sample workbook from the configured workbook class or sheets.
     *
     * @return the creation result (workbook and optional password)
     * @throws PxlException if both or neither of the workbook/sheet forms are specified, or if workbook creation fails
     */
    @Override
    protected Built build()
            throws PxlException {

        if (Objects.nonNull(workbookClass) && PxlCollectionUtils.isNotEmpty(sheetNames)) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.BUILDER_WORKBOOK_SHEET_EXCLUSIVE, "workbook(Class)"));
        }

        if (Objects.isNull(workbookClass) && PxlCollectionUtils.isEmpty(sheetNames)) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.BUILDER_WORKBOOK_SHEET_REQUIRED, "workbook(Class)"));
        }

        PxlExportWorkbookMeta workbookMeta = null;
        boolean success = false;

        try {
            final Workbook workbook;
            if (Objects.nonNull(workbookClass)) {
                workbookMeta = PxlExportWorkbookMeta.makeExportWorkbookMeta(workbookClass, option);
                workbook = PxlExcelExporter.buildSampleWorkbook(workbookClass, workbookMeta);
            } else {
                workbookMeta = PxlExportWorkbookMeta.makeExportWorkbookMeta(null, option);
                workbook = PxlExcelExporter.buildSampleWorkbook(sheetNames, rowClasses, workbookMeta);
            }
            success = true;

            return new Built(workbook, workbookMeta.getExportPassword());
        } catch (PxlException e) {
            throw e;
        } catch (Exception e) {
            throw new PxlSystemException(e);
        } finally {
            if (!success && Objects.nonNull(workbookMeta)) {
                PxlWorkbookUtils.closeWorkbook(workbookMeta.getWorkbook());
            }
        }
    }

}
