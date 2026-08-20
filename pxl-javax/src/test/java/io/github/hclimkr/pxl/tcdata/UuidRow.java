package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.util.List;
import java.util.UUID;

/**
 * DTO for verifying UUID column binding in both directions.
 * Each column isolates one aspect: the plain value, a collection of values, the untrimmed parse, the export mask,
 * and the import-time uniqueness check. The columns are ordered explicitly, since field declaration order does not
 * decide column order.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class UuidRow {

    @PxlColumn(name = "Id", exportOrder = "1", exportSample = "123e4567-e89b-12d3-a456-426614174000")
    private UUID id;

    // Collection of UUID elements, joined/split with the default separator (";").
    @PxlColumn(name = "Ids", exportOrder = "2")
    private List<UUID> ids;

    // importTrim disabled, so a padded value reaches the codec exactly as it was written.
    @PxlColumn(name = "Exact", exportOrder = "3", importTrim = false)
    private UUID exact;

    // Masks every hexadecimal digit, leaving the hyphens of the canonical form in place.
    @PxlColumn(name = "Masked", exportOrder = "4", exportMasking = "[0-9a-f]")
    private UUID masked;

    // Duplicate values in this column fail the import-time uniqueness check.
    @PxlColumn(name = "Unique", exportOrder = "5", importUnique = true)
    private UUID unique;

}
