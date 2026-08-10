package io.github.hclimkr.pxl.type;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.annotation.PxlWorkbook;

/**
 * A yes/no setting that can also say nothing at all.
 *
 * <p>An annotation element cannot hold {@code null}, so a {@code boolean} one has no way to mean "not specified":
 * its {@code false} is indistinguishable from a deliberate {@code false}. That is only a problem where a setting
 * cascades through several annotation levels, because the inner level then has no way to leave the outer one
 * alone — and no way to turn it back off either, since {@code false} would be read as silence.</p>
 *
 * <p>Where a {@link String} element says "not specified" with an empty value and a {@code char} one with NUL, a
 * three-valued element says it with {@link #UNSPECIFIED}. Option classes keep using a boxed {@link Boolean} whose
 * {@code null} already carries that meaning.</p>
 *
 * @see PxlSheet#exportCsvBom()
 * @see PxlWorkbook#exportCsvBom()
 */
public enum PxlOptionalBoolean {

    /**
     * Not specified at this level: the next level of the cascade decides.
     */
    UNSPECIFIED,

    /**
     * Specified as true.
     */
    TRUE,

    /**
     * Specified as false, overriding an outer level that said true.
     */
    FALSE;

    /**
     * Answers whether this level decides the value rather than deferring to the next one.
     *
     * @return {@code true} unless this is {@link #UNSPECIFIED}
     */
    public boolean isSpecified() {

        return this != UNSPECIFIED;
    }

    /**
     * Returns this value as a plain boolean.
     *
     * <p>{@link #UNSPECIFIED} answers {@code false}, so call {@link #isSpecified()} first wherever the difference
     * between "not specified" and "specified as false" matters — which is everywhere the cascade is walked.</p>
     *
     * @return {@code true} only for {@link #TRUE}
     */
    public boolean toBoolean() {

        return this == TRUE;
    }

}
