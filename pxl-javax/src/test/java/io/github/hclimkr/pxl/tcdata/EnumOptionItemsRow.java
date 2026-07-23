package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO combining an enum column with an explicit {@code exportOptionItems} list and a dropdown style,
 * exercising the "enum + SET/SORTED_SET + option items" branches of the dropdown builder
 * (SET keeps the given order, SORTED_SET sorts the items).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class EnumOptionItemsRow {

    @PxlColumn(name = "GradeSet",
            exportEnumDropDownListStyle = PxlColumn.ExportEnumDropDownListStyle.SET,
            exportOptionItems = {"B", "A", "C"})
    private Grade gradeSet;

    @PxlColumn(name = "GradeSorted",
            exportEnumDropDownListStyle = PxlColumn.ExportEnumDropDownListStyle.SORTED_SET,
            exportOptionItems = {"C", "A", "B"})
    private Grade gradeSorted;

}
