package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.constraint.PxlByteSize;
import lombok.*;

/**
 * DTO for verifying the public constraint annotation @PxlByteSize.
 * <ul>
 *   <li>Code: max 5 bytes (UTF-8)</li>
 *   <li>Name: min 4 bytes (UTF-8) — for lower-bound verification</li>
 *   <li>Label: max 4 bytes (EUC-KR) — for verifying charset-based byte counting</li>
 * </ul>
 * Each column is independent; if absent from the sheet or null, @PxlByteSize passes as valid (null allowed).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ByteSizeRow {

    @PxlByteSize(max = 5)
    @PxlColumn(name = "Code")
    private String code;

    @PxlByteSize(min = 4)
    @PxlColumn(name = "Name")
    private String name;

    @PxlByteSize(max = 4, charset = "EUC-KR")
    @PxlColumn(name = "Label")
    private String label;

}
