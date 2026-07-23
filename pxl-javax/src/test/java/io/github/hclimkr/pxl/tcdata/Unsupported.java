package io.github.hclimkr.pxl.tcdata;

/**
 * A type PXL does not support (no converter, no String constructor, no toString override).
 * A column of this type must fail-fast during meta build.
 */
public class Unsupported {

    private int value;

}
