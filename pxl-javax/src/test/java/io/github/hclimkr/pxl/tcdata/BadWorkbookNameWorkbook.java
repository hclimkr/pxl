package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

/**
 * A @PxlWorkbook-style DTO whose @PxlWorkbookName field is (incorrectly) not a String,
 * used to verify PxlWorkbookSupport.validateWorkbookNameFieldType rejects a non-String name field.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class BadWorkbookNameWorkbook {

    @PxlWorkbookName
    private Integer workbookName;

}
