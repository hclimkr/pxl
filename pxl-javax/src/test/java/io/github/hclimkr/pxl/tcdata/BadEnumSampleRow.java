package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO where the exportSample of an enum column does not match any constant of that enum (issue N3).
 * For verifying that sample export fails early via fail-fast (PxlArgumentException) instead of crashing during cell writing.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class BadEnumSampleRow {

    @PxlColumn(name = "Grade", exportSample = "N/A")   // Grade(A/B/C/F) has no N/A constant
    private Grade grade;

}
