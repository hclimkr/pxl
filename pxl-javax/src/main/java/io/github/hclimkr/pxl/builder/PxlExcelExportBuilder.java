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

import javax.validation.Validator;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Objects;

/**
 * Builder that exports data to an Excel workbook/file/stream. Created via {@link io.github.hclimkr.pxl.Pxl#exportExcel()}.
 *
 * <p>There are two usage forms.</p>
 * <ul>
 *   <li>{@code @PxlWorkbook} object form: {@link #workbook(Object)}</li>
 *   <li>Sheet form (multiple calls for multiple sheets): {@link #sheet(String, Collection, Class)}</li>
 * </ul>
 *
 * <p>The two forms are mutually exclusive — specifying both {@code workbook(...)} and {@code sheet(...)} throws an exception.</p>
 *
 * <p>The terminal methods (returning a workbook / file / stream) and resource handling are provided by {@link PxlAbstractExportBuilder}.</p>
 *
 * <p>Example: {@code pxl.exportExcel().sheet("Users", users, User.class).override(opt).toFile(file);}</p>
 */
public final class PxlExcelExportBuilder extends PxlAbstractExportBuilder {

    private final Validator validator;

    private Object workbookObject;

    private final List<String> sheetNames = new ArrayList<>();
    private final List<Collection<?>> sheetObjects = new ArrayList<>();
    private final List<Class<?>> rowClasses = new ArrayList<>();

    /**
     * Creates an Excel export builder with the given validator.
     *
     * @param validator the bean-validation validator, or {@code null} when bean validation is disabled
     */
    public PxlExcelExportBuilder(final Validator validator) {

        this.validator = validator;
    }

    /**
     * Exports from a workbook object annotated with {@code @PxlWorkbook}. (Mutually exclusive with the sheet form)
     *
     * @param workbookObject the {@code @PxlWorkbook}-annotated source object
     * @return this builder
     * @throws PxlNullPointerException if {@code workbookObject} is {@code null}
     */
    public PxlExcelExportBuilder workbook(final Object workbookObject)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(workbookObject, "workbookObject");

        this.workbookObject = workbookObject;
        return this;
    }

    /**
     * Adds a sheet. Calling multiple times produces multiple sheets. (Mutually exclusive with the workbook object form)
     *
     * @param sheetName the sheet name; must not be blank
     * @param rows      the row objects for this sheet
     * @param rowClass  the row class
     * @param <T>       the row type
     * @return this builder
     * @throws PxlNullPointerException if {@code sheetName}, {@code rows}, or {@code rowClass} is {@code null}
     * @throws PxlArgumentException    if {@code sheetName} is blank
     */
    public <T> PxlExcelExportBuilder sheet(final String sheetName,
                                           final Collection<T> rows,
                                           final Class<T> rowClass)
            throws PxlNullPointerException, PxlArgumentException {

        PxlAssertSupport.notBlank(sheetName, "sheetName");
        PxlAssertSupport.notNull(rows, "rows");
        PxlAssertSupport.notNull(rowClass, "rowClass");

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
    public PxlExcelExportBuilder override(@Nullable final PxlExportWorkbookOption option) {

        this.option = option;
        return this;
    }

    /**
     * Builds the Excel workbook from the configured workbook object or sheets.
     *
     * @return the creation result (workbook and optional password)
     * @throws PxlException if both or neither of the workbook/sheet forms are specified, or if workbook creation fails
     */
    @Override
    protected Built build()
            throws PxlException {

        if (Objects.nonNull(workbookObject) && PxlCollectionUtils.isNotEmpty(sheetNames)) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.BUILDER_WORKBOOK_SHEET_EXCLUSIVE, "workbook(Object)"));
        }

        if (Objects.isNull(workbookObject) && PxlCollectionUtils.isEmpty(sheetNames)) {
            throw new PxlArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.BUILDER_WORKBOOK_SHEET_REQUIRED, "workbook(Object)"));
        }

        PxlExportWorkbookMeta workbookMeta = null;
        boolean success = false;

        try {
            final Workbook workbook;
            if (Objects.nonNull(workbookObject)) {
                workbookMeta = PxlExportWorkbookMeta.makeExportWorkbookMeta(workbookObject.getClass(), option);
                workbook = PxlExcelExporter.buildWorkbook(workbookObject, workbookMeta, validator);
            } else {
                workbookMeta = PxlExportWorkbookMeta.makeExportWorkbookMeta(null, option);
                workbook = PxlExcelExporter.buildWorkbook(sheetNames, sheetObjects, rowClasses, workbookMeta, validator);
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
