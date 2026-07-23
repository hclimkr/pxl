package io.github.hclimkr.pxl.internal.support;

import io.github.hclimkr.pxl.internal.meta.PxlExportColumnMeta;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.Arrays;

/**
 * Column-related utilities.
 */
public final class PxlColumnSupport {

    /**
     * Prevents instantiation.
     */
    private PxlColumnSupport() {

        throw new AssertionError("no instances of this class");
    }

    /**
     * Adds a dropdown (list) data validation over the given cell range.
     * <p>
     * When the items are short enough and contain no comma, an inline explicit-list constraint is used. Otherwise (an item
     * contains a comma, or the explicit-list formula would exceed Excel's 255-character limit) the items are written to a
     * hidden helper sheet and referenced through a defined name (formula-list constraint). Helper sheet and defined names
     * are made unique within the workbook to avoid collisions when the same (sheet, column) is reached twice.
     * <p>
     * No-ops when any of {@code sheet}, {@code columnMeta}, {@code cellRangeAddressList} is {@code null} or {@code itemStrings} is empty.
     *
     * @param sheet                the sheet the validation is applied to
     * @param columnMeta           the export column metadata (supplies the sheet/column index for naming)
     * @param cellRangeAddressList the cell range the dropdown applies to
     * @param itemStrings          the dropdown item values
     */
    public static void setDropdownList(final Sheet sheet,
                                       final PxlExportColumnMeta columnMeta,
                                       final CellRangeAddressList cellRangeAddressList,
                                       final String[] itemStrings) {

        if (ObjectUtils.anyNull(sheet, columnMeta, cellRangeAddressList) || ArrayUtils.isEmpty(itemStrings)) {
            return;
        }

        final DataValidationHelper validationHelper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint;

        final boolean anyItemHasComma = Arrays.stream(itemStrings).anyMatch(item -> StringUtils.contains(item, ","));

        final String formula = StringUtils.join(itemStrings, ",");
        // The explicit-list formula is subject to the 255-char limit, including the enclosing quotes (2 chars) and the doubling of inner quotes.
        final int explicitListFormulaLength = 2 + StringUtils.length(StringUtils.replace(formula, "\"", "\"\""));

        if (anyItemHasComma || explicitListFormulaLength > 255) {
            // Excel limitation: a formula length cannot exceed 255 characters.
            // In this case the items must be written to a hidden sheet and a formula referencing that range must be used.

            final Workbook workbook = sheet.getWorkbook();
            final int sheetIndex = workbook.getSheetIndex(sheet);
            final int columnIndex = columnMeta.getActualExportColumnIndex();

            // If the name were deterministic, reaching the same (sheet,column) twice in the same workbook would make
            // createSheet/createName fail with IllegalArgumentException("already contains…"), so make a name unique within the workbook to avoid the collision.
            final String desiredSheetName = "dropdown_sheet_s" + sheetIndex + "_c" + columnIndex;
            final String dropdownSheetName = PxlWorkbookSupport.makeUniqueSafeSheetName(workbook, desiredSheetName);
            final Sheet dropdownSheet = workbook.createSheet(dropdownSheetName);

            for (int i = 0; i < itemStrings.length; i++) {
                final Row row = dropdownSheet.createRow(i);
                final Cell cell = row.createCell(0);
                cell.setCellValue(itemStrings[i]);
            }

            final String desiredNamedRangeName = "dropdown_named_range_s" + sheetIndex + "_c" + columnIndex;
            final String dropdownNamedRangeName = PxlWorkbookSupport.makeUniqueDefinedName(workbook, desiredNamedRangeName);
            final Name dropdownNamedRange = workbook.createName();
            dropdownNamedRange.setNameName(dropdownNamedRangeName);

            // The uniquifying suffix (e.g. " (2)") may introduce spaces/special characters into the sheet name, so always quote the sheet name in the formula.
            final String quotedDropdownSheetName = "'" + dropdownSheetName.replace("'", "''") + "'";
            dropdownNamedRange.setRefersToFormula(quotedDropdownSheetName + "!$A$1:$A$" + itemStrings.length);

            workbook.setSheetHidden(workbook.getSheetIndex(dropdownSheet), true);

            constraint = validationHelper.createFormulaListConstraint(dropdownNamedRangeName);
        } else {
            constraint = validationHelper.createExplicitListConstraint(itemStrings);
        }

        final DataValidation dataValidation = validationHelper.createValidation(constraint, cellRangeAddressList);

        dataValidation.setEmptyCellAllowed(false);
        if (dataValidation instanceof HSSFDataValidation) {
            dataValidation.setSuppressDropDownArrow(false);
        } else {
            dataValidation.setSuppressDropDownArrow(true);
        }
        dataValidation.setShowPromptBox(false);
        // dataValidation.createErrorBox("error title", "error text");
        dataValidation.setShowErrorBox(true);
        // dataValidation.createPromptBox("prompt title", "prompt text");
        // dataValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        sheet.addValidationData(dataValidation);
    }

}
