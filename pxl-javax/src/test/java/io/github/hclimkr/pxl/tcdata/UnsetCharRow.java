package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;
import lombok.*;

/**
 * DTO for verifying how a primitive {@code char} that was never set is exported. Such a field holds {@code (char) 0} -
 * the only way the type has of saying "no value", since a {@code char} cannot be {@code null} - and export renders it
 * as the column's export-null string rather than writing the NUL character itself into the cell. The set column is the
 * control, and the {@link Character} column is the counterpart where {@code null} is what absence already looks like.
 * The columns are ordered explicitly, since field declaration order does not decide column order.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class UnsetCharRow {

    // Left at (char) 0, so the default exportNullString ("") decides the cell.
    @PxlColumn(name = "UnsetChar", exportOrder = "1")
    private char unsetChar;

    // Same value, but with a custom exportNullString, which shows that the value really took the absent-value path.
    @PxlColumn(name = "UnsetCharDash", exportOrder = "2", exportNullString = "-")
    private char unsetCharDash;

    // The control: a char that was given a value is written as that character.
    @PxlColumn(name = "SetChar", exportOrder = "3")
    private char setChar;

    // The boxed counterpart, left null.
    @PxlColumn(name = "UnsetWrapChar", exportOrder = "4")
    private Character unsetWrapChar;

}
