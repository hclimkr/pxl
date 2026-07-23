package io.github.hclimkr.pxl.internal.core;

import io.github.hclimkr.pxl.annotation.PxlRowIndex;
import io.github.hclimkr.pxl.exception.PxlNullPointerException;
import io.github.hclimkr.pxl.exception.PxlReflectionException;
import io.github.hclimkr.pxl.internal.constraint.Nullable;
import io.github.hclimkr.pxl.internal.meta.PxlImportColumnMeta;
import io.github.hclimkr.pxl.internal.support.PxlAssertSupport;
import io.github.hclimkr.pxl.internal.support.PxlReflectionSupport;
import io.github.hclimkr.pxl.internal.support.PxlWorkbookSupport;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Objects;

/**
 * Excel import common routine
 */
abstract class PxlAbstractImporter extends PxlAbstractBinder {

    /**
     * Assigns the workbook name to the {@code @PxlWorkbookName}-annotated field of the workbook object, if any.
     * Silently ignores objects without such a field and swallows any reflection failure.
     *
     * @param workbookObject the workbook object whose name field is populated
     * @param workbookName   the name to set (may be {@code null})
     * @throws PxlNullPointerException if {@code workbookObject} is {@code null}
     */
    protected static void setWorkbookNameToWorkbookObject(final Object workbookObject,
                                                          @Nullable final String workbookName)
            throws PxlNullPointerException {

        PxlAssertSupport.notNull(workbookObject, "workbookObject");

        final Field workbookNameField = PxlWorkbookSupport.getWorkbookNameField(workbookObject.getClass());
        if (Objects.nonNull(workbookNameField)) {
            try {
                PxlReflectionSupport.setFieldValue(workbookNameField, workbookObject, workbookName);
            } catch (Exception ignored) {
            }
        }
    }

    /**
     * Finds the first field of the row class (including inherited fields) annotated with {@code @PxlRowIndex}.
     *
     * @param rowClass the row class to inspect
     * @return the {@code @PxlRowIndex}-annotated field, or {@code null} if none exists
     */
    protected static Field getRowIndexField(final Class<?> rowClass) {

//        return Arrays.stream(rowClass.getDeclaredFields())
//                .filter(o -> Objects.nonNull(o.getAnnotation(PxlRowIndex.class)))
//                .findFirst()
//                .orElse(null);
        return PxlReflectionSupport.getAllFields(rowClass).stream()
                .filter(o -> Objects.nonNull(o.getAnnotation(PxlRowIndex.class)))
                .findFirst()
                .orElse(null);
    }

    /**
     * Returns whether the row can be skipped because it carries no data.
     * A row is ignorable when every import-enabled column field of the row object is {@code null}.
     *
     * @param columnMetas the per-column import metadata
     * @param rowObject   the populated row object to inspect
     * @return {@code true} if all import-enabled column fields are {@code null}; {@code false} otherwise
     * @throws PxlReflectionException if a column field value cannot be read
     */
    protected static boolean isIgnorableRow(final List<PxlImportColumnMeta> columnMetas,
                                            final Object rowObject)
            throws PxlReflectionException {

        // If the row object holds no values at all, it is an empty row and is skipped.
        for (final PxlImportColumnMeta columnMeta : columnMetas) {
            if (!columnMeta.isImportEnabled()) {
                continue;
            }

            final Field columnField = columnMeta.getColumnField();

            if (Objects.nonNull(PxlReflectionSupport.getFieldValue(columnField, rowObject))) {
                return false;
            }
        }

        return true;
    }

}
