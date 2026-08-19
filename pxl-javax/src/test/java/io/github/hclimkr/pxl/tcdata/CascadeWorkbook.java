package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import javax.validation.Valid;
import java.util.List;

/**
 * Workbook form whose sheet field carries {@code @Valid}, so bean validation cascades from the workbook object
 * into every row. In this form each row is therefore validated twice - once through the cascade, once by the
 * per-row/per-collection pass the binder runs on its own.
 * <p>
 * {@link NoCascadeWorkbook} declares the same sheet without {@code @Valid} and is the control for it.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class CascadeWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @Valid                          // <- the cascade switch
    @PxlSheet(name = "Rows")
    private List<CountingRow> rows;

}
