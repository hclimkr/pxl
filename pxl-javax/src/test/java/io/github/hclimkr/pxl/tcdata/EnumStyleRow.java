package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * A DTO for verifying exportEnumDropDownListStyle combinations.
 * none means no dropdown is created (NONE), sorted means a sorted dropdown (SORTED_SET).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class EnumStyleRow {

    @PxlColumn(name = "GradeNone", exportEnumDropDownListStyle = PxlColumn.ExportEnumDropDownListStyle.NONE)
    private Grade gradeNone;

    @PxlColumn(name = "GradeSorted", exportEnumDropDownListStyle = PxlColumn.ExportEnumDropDownListStyle.SORTED_SET)
    private Grade gradeSorted;

}
