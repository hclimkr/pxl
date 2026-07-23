package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO binding NUMERIC cells (a plain integer, a date-formatted cell, a large integer) to {@code String} fields.
 * <p>
 * Verifies that reading numeric cells into {@code String} via the streaming reader renders them with the cell's
 * display format through POI's {@code DataFormatter} — the same as non-streaming — now that the former
 * {@code StreamingCell}/{@code NumberToTextConverter} special-case has been removed from {@code PxlStringCodec}.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class StringFromNumericRow {

    // Plain General-format numeric cell -> "123"
    @PxlColumn(name = "PlainInt")
    private String plainInt;

    // Date-formatted numeric cell (yyyy-mm-dd) -> "2020-01-15" (display format, not the raw Excel serial)
    @PxlColumn(name = "DateFormatted")
    private String dateFormatted;

    // Large integer numeric cell -> "2012000046" (General format, without exponent notation)
    @PxlColumn(name = "LargeInt")
    private String largeInt;

}
