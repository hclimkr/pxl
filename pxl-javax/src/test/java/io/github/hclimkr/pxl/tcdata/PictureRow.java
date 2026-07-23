package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.util.List;

/**
 * DTO for verifying export of a string (image URL) as an actual image cell (exportStringAsPicture).
 * <p>
 * Covers both paths: a single String column and a List&lt;String&gt; column (multiple images in one cell).
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class PictureRow {

    // Single image
    @PxlColumn(name = "Photo", exportStringAsPicture = true)
    private String photo;

    // Multiple images in one cell
    @PxlColumn(name = "Gallery", exportStringAsPicture = true)
    private List<String> gallery;

}
