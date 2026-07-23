package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import javax.validation.constraints.NotBlank;
import javax.validation.constraints.NotNull;

/**
 * Row DTO with Bean Validation constraints. Verifies that a missing required value throws an exception on export/import.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ValidatedRow {

    @NotBlank(message = "'Name' must not be blank.")
    @PxlColumn(name = "Name")
    private String name;

    @NotNull(message = "'Age' must not be null.")
    @PxlColumn(name = "Age")
    private Integer age;

}
