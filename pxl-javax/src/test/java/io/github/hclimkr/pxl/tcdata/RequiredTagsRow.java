package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import javax.validation.constraints.NotEmpty;
import java.util.List;

/**
 * Row DTO with a {@code @NotEmpty} collection field. Verifies empty cell → {@code null} collection → {@code @NotEmpty} violation.
 * An unconstrained {@code name} field is included alongside to avoid empty-row skipping ({@code isIgnorableRow}).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class RequiredTagsRow {

    @PxlColumn(name = "Name")
    private String name;

    @NotEmpty(message = "'Tags' must not be empty.")
    @PxlColumn(name = "Tags")
    private List<String> tags;

}
