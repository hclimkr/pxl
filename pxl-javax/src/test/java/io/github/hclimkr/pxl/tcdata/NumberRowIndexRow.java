package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import io.github.hclimkr.pxl.annotation.PxlRowIndex;
import lombok.*;

/**
 * DTO with a {@link Number} @PxlRowIndex field, exercising the importer's fallback ({@code cast}) branch of
 * row-index injection (the type is neither byte/short/int/long nor their wrapper classes Byte/Short/Integer/Long, but an int is still
 * assignable to it, so the cast succeeds).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class NumberRowIndexRow {

    @PxlRowIndex
    private Number rowIndex;

    @PxlColumn(name = "Name")
    private String name;

}
