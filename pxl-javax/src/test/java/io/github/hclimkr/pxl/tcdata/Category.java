package io.github.hclimkr.pxl.tcdata;

/**
 * Enum that overrides toString.
 * Export goes through toString (label); import round-trips via label matching (whitespace-insensitive/case-insensitive).
 */
public enum Category {

    ELECTRONICS("Electronics"),
    FOOD("Food & Beverage"),
    CLOTHING("Clothing"),
    ;

    private final String label;

    Category(final String label) {
        this.label = label;
    }

    @Override
    public String toString() {
        return this.label;
    }

}
