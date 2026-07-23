package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.time.OffsetDateTime;
import java.time.ZonedDateTime;

/**
 * Verifies offset preservation for strings containing an ISO offset.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ZonedRow {

    @PxlColumn(name = "Zoned")
    private ZonedDateTime zoned;

    @PxlColumn(name = "Offset")
    private OffsetDateTime offset;

}
