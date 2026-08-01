package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * Row for verifying i18n. The column name becomes the key in the message bundle, so the header is translated.
 * (messages.properties: staff.column.role=Role, staff.column.fullName=Full Name)
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class I18nRow {

    @PxlColumn(name = "staff.column.role")
    private String role;

    @PxlColumn(name = "staff.column.fullName")
    private String fullName;

}
