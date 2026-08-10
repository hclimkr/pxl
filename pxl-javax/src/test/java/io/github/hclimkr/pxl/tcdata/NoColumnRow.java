package io.github.hclimkr.pxl.tcdata;

import lombok.*;

/**
 * DTO that binds no column at all, leaving an export with nothing to write.
 * <p>
 * Distinct from {@link NoSampleColumnRow}, whose columns exist but opt out of the sample: there the column metadata
 * is built and simply left unmapped, whereas here there is no column metadata to begin with.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class NoColumnRow {

    private String notAColumn;

}
