package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.util.List;

/**
 * Row for verifying content i18n of exportSample and exportOptionItems.
 * (messages.properties: staff.column.role=Role, .roles=Roles, .grades=Grades, .grade=Grade, .tags=Tags,
 * staff.role.admin=Administrator, staff.role.user=User, staff.grade.a=A, staff.grade.b=B, 1234=9999;
 * the "count" column name is absent from the bundle on purpose)
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class I18nSampleRow {

    @PxlColumn(name = "staff.column.role", exportSample = "staff.role.admin", exportOptionItems = {"staff.role.admin", "staff.role.user"})
    private String role;

    // A Collection sample holds one key per element, separated by the collection separator (default ";").
    @PxlColumn(name = "staff.column.roles", exportSample = "staff.role.admin;staff.role.user")
    private List<String> roles;

    // Collection<Enum>: each element is translated first, then parsed back into its constant.
    @PxlColumn(name = "staff.column.grades", exportSample = "staff.grade.a;staff.grade.b")
    private List<Grade> grades;

    // An enum column writes the canonical constant, not the translation, so its option items stay verbatim.
    @PxlColumn(name = "staff.column.grade", exportSample = "staff.grade.a", exportOptionItems = {"staff.grade.a", "staff.grade.b"})
    private Grade grade;

    // The separator belongs to the column, not to the bundle value, so a custom one still splits the sample.
    @PxlColumn(name = "staff.column.tags", exportSample = "staff.role.admin::staff.role.user", exportCollectionSeparator = "::")
    private List<String> tags;

    // A non-String/enum column keeps its sample verbatim: "1234" is a bundle key (1234=9999) that must not be
    // applied, and "count" is deliberately absent from the bundle so the header passes through untranslated.
    @PxlColumn(name = "count", exportSample = "1234")
    private Integer count;

}
