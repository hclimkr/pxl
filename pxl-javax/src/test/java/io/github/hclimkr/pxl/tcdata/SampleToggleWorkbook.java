package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import java.util.List;

/**
 * Workbook for verifying sheet-level exportSampleEnabled.
 * withSample is included in the sample; noSample (exportSampleEnabled=false) is excluded from the sample.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class SampleToggleWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "WithSample")
    private List<Employee> withSample;

    @PxlSheet(name = "NoSample", exportSampleEnabled = false)
    private List<Employee> noSample;

}
