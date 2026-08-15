package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.util.Date;

/**
 * DTO where the exportSample of a {@link Date} column carries text beyond what the pattern reads.
 * For verifying that the sample export rejects it instead of writing the date it managed to read (issue M2).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class BadJavaDateSampleRow {

    @PxlColumn(name = "D", pattern = "yyyy-MM-dd", exportSample = "2020-01-15 xxx")   // the pattern stops at the space
    private Date d;

}
