package io.github.hclimkr.pxl.tcdata;

import lombok.*;

import javax.validation.Valid;
import java.util.List;

/**
 * Plain nested object held by {@link ConstrainedCascadeWorkbook} through a field that is not a sheet. It carries
 * no Pxl annotation, so the binder ignores it entirely and only bean validation ever looks at it - which is what
 * makes it the control for "the sheet cascade is skipped, every other cascade is not".
 * <p>
 * Its own {@code @Valid} list goes one level deeper: rows reached below the root are not rows the binder walks
 * one by one, so that cascade has to survive as well even though the element type is the very same row class.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class CascadeMeta {

    @CountedNotBlank(message = "'Label' must not be blank.")
    private String label;

    @Valid
    private List<CountingRow> nestedRows;

}
