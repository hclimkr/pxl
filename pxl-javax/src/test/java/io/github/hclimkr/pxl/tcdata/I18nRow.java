package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * Row for verifying i18n. The column name becomes the key in the message bundle, so the header is translated.
 * (messages.properties: role=Role, fullname=Full Name)
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class I18nRow {

    @PxlColumn(name = "role")
    private String role;

    @PxlColumn(name = "fullname")
    private String fullName;

}
