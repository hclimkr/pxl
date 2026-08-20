package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.util.UUID;

/**
 * DTO whose UUID column carries an exportSample that is not a canonical UUID.
 * For verifying that sample export fails instead of writing a value the import side would refuse. The sample is one
 * that {@code UUID.fromString} itself accepts - it counts the hyphen-separated groups but not their digits - so this
 * also pins that the codec, not the JDK, decides what a UUID column accepts.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class BadUuidSampleRow {

    @PxlColumn(name = "Id", exportSample = "1-1-1-1-1")
    private UUID id;

}
