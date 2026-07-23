package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlExportConverter;
import io.github.hclimkr.pxl.annotation.PxlImportConverter;
import lombok.*;

/**
 * Custom object type.
 * Import round-trips via @PxlImportConverter (static method), export via @PxlExportConverter (instance method).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class Money {

    private String currency;

    private long amount;

    // String -> Money (import path)
    @PxlImportConverter
    public static Money fromExportString(final String value) {
        final String[] parts = value.trim().split("\\s+");
        return new Money(parts[0], Long.parseLong(parts[1]));
    }

    // Money -> String (export path)
    @PxlExportConverter
    public String toExportString() {
        return this.currency + " " + this.amount;
    }

}
