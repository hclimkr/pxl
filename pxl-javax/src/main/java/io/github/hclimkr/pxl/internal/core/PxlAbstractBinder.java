package io.github.hclimkr.pxl.internal.core;

import io.github.hclimkr.pxl.exception.PxlReflectionException;
import io.github.hclimkr.pxl.exception.PxlValidationException;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;
import io.github.hclimkr.pxl.internal.meta.PxlImportColumnMeta;
import io.github.hclimkr.pxl.internal.support.PxlReflectionSupport;
import io.github.hclimkr.pxl.util.PxlCollectionUtils;
import org.apache.commons.lang3.ObjectUtils;

import javax.validation.ConstraintViolation;
import javax.validation.Validator;
import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Common base for the export and import binders. Provides shared bean-validation
 * ({@code javax.validation}) of workbook/sheet/row objects and per-column uniqueness checks.
 */
abstract class PxlAbstractBinder {

    /**
     * Validates a single object against its bean-validation constraints and throws when any is violated.
     * Does nothing when the validator or the object is {@code null}.
     *
     * @param validator the bean validator to apply; validation is skipped when {@code null}
     * @param object    the object to validate; validation is skipped when {@code null}
     * @param sheetName the sheet name reported in the thrown exception (may be {@code null})
     * @param rowIndex  the row index reported in the thrown exception (may be {@code null})
     * @throws PxlValidationException if one or more constraints are violated; the message concatenates the distinct violation messages
     */
    protected static void validateBeanConstraints(final Validator validator,
                                                  final Object object,
                                                  final String sheetName,
                                                  final Integer rowIndex)
            throws PxlValidationException {

        if (ObjectUtils.anyNull(validator, object)) {
            return;
        }

        final Set<ConstraintViolation<Object>> violations = validator.validate(object);
        if (PxlCollectionUtils.isNotEmpty(violations)) {
            final String errMsg = violations.stream()
                    .filter(Objects::nonNull)
                    .map(ConstraintViolation::getMessage)
                    .distinct()
                    .collect(Collectors.joining(", "));

            throw new PxlValidationException(sheetName, rowIndex, null, null, errMsg);
        }
    }

    /**
     * Validates each element of the given collection individually against its bean-validation constraints.
     * Does nothing when the validator is {@code null} or the collection is empty.
     *
     * @param validator the bean validator to apply; validation is skipped when {@code null}
     * @param objects   the objects to validate; each non-null element is validated separately
     * @param sheetName the sheet name reported in the thrown exception (may be {@code null})
     * @param rowIndex  the row index reported in the thrown exception (may be {@code null})
     * @throws PxlValidationException if any element violates a constraint
     */
    protected static void validateBeanConstraints(final Validator validator,
                                                  final Collection<?> objects,
                                                  final String sheetName,
                                                  final Integer rowIndex)
            throws PxlValidationException {

        if (Objects.isNull(validator) || PxlCollectionUtils.isEmpty(objects)) {
            return;
        }

        // TODO: look into a way to validate rowObjects all at once.
        // final Set<ConstraintViolation<Collection<Object>>> violations = validator.validate(objects);

        for (final Object object : objects) {
            validateBeanConstraints(validator, object, sheetName, rowIndex);
        }
    }

    /**
     * Checks that columns marked unique contain no duplicate values across the given row objects.
     * For each import-enabled column whose index was resolved and whose {@code importUnique} flag is set,
     * the field values of all rows are collected and inspected for duplicates.
     *
     * @param objects     the row objects whose column values are checked
     * @param columnMetas the per-column import metadata; only enabled, resolved, unique-marked columns are checked
     * @param sheetName   the sheet name reported in the thrown exception (may be {@code null})
     * @throws PxlReflectionException if a column field value cannot be read
     * @throws PxlValidationException if a unique-marked column contains duplicate values
     */
    protected static void validateDataUniqueness(final Collection<?> objects,
                                                 final List<PxlImportColumnMeta> columnMetas,
                                                 final String sheetName)
            throws PxlReflectionException, PxlValidationException {

        if (PxlCollectionUtils.isEmpty(objects)) {
            return;
        }

        for (final PxlImportColumnMeta columnMeta : columnMetas) {
            if (!columnMeta.isImportEnabled()) {
                continue;
            }

            final int importColumnIndex = columnMeta.getActualImportColumnIndex();
            if (importColumnIndex < 0) {
                continue;
            }

            if (!columnMeta.isImportUnique()) {
                continue;
            }

            final Field columnField = columnMeta.getColumnField();
            final List<Object> fieldValues = new ArrayList<>(PxlCollectionUtils.size(objects));
            for (final Object rowObject : objects) {
                if (Objects.nonNull(rowObject)) {
                    fieldValues.add(PxlReflectionSupport.getFieldValue(columnField, rowObject));
                }
            }
            final Set<Object> duplicates = PxlCollectionUtils.findDuplicates(fieldValues);

            if (PxlCollectionUtils.isNotEmpty(duplicates)) {
                final String duplicateValues = duplicates.stream()
                        .map(Object::toString)
                        .collect(Collectors.joining(", "));

                throw new PxlValidationException(sheetName, null, columnMeta.getActualImportColumnName(), null,
                        PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.CORE_IMPORT_COLUMN_VALUE_DUPLICATE, duplicateValues));
            }
        }
    }

}
