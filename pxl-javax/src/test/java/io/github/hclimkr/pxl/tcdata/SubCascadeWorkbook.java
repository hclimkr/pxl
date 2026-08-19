package io.github.hclimkr.pxl.tcdata;

/**
 * Subclass adding nothing to {@link CascadeWorkbook}, so its {@code @Valid @PxlSheet} field is inherited rather
 * than declared.
 * <p>
 * The sheet-cascade resolver caches its decision per class, keyed by the class actually being validated - here the
 * subclass - so this fixture is what proves the scan walks the superclass chain instead of only the declared
 * fields. Carries no Lombok annotation because it declares no field of its own; the parent's accessors are
 * inherited as they are.
 */
public class SubCascadeWorkbook extends CascadeWorkbook {

}
