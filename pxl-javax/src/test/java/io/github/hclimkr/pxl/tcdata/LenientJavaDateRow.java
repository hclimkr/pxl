package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

import java.util.Date;

/**
 * DTO characterizing the behavior where java.util.Date + a custom importPattern is lenient (missing setLenient(false)),
 * so invalid dates silently roll over. Unlike the built-in Date reader and java.time (strict), only Date + custom pattern is lenient.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class LenientJavaDateRow {

    @PxlColumn(name = "D", importPattern = "yyyy-MM-dd")
    private Date d;

}
