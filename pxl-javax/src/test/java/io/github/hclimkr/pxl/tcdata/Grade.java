package io.github.hclimkr.pxl.tcdata;

/**
 * Simple enum that does not override toString.
 * Round-trips via name() on export and name() matching on import.
 */
public enum Grade {

    A,
    B,
    C,
    F,

}
