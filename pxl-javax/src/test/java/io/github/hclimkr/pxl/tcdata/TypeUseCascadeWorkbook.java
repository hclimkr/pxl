package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbookName;
import lombok.*;

import javax.validation.Valid;
import java.util.List;

/**
 * Same cascade as {@link CascadeWorkbook}, declared the other way Bean Validation 2.0 allows: as a container
 * element constraint ({@code List<@Valid Row>}) rather than on the field.
 * <p>
 * Both forms are cascade switches, so both have to be caught by the sheet-cascade resolver; this fixture is what
 * proves the container element form is not a hole in it.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class TypeUseCascadeWorkbook {

    @PxlWorkbookName
    private String workbookName;

    @PxlSheet(name = "Rows")
    private List<@Valid CountingRow> rows;

}
