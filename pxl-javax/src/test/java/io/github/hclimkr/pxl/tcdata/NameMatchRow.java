package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * Verifies that column-name matching ignores whitespace and is case-sensitive. ("FullName" vs "Full Name"/"fullname")
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class NameMatchRow {

    @PxlColumn(name = "FullName")
    private String name;

    // Auxiliary column that always matches (so the sheet has at least one valid column)
    @PxlColumn(name = "Id")
    private String id;

}
