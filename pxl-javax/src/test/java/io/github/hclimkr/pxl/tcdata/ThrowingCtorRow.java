package io.github.hclimkr.pxl.tcdata;

import io.github.hclimkr.pxl.annotation.PxlColumn;

/**
 * DTO whose no-arg constructor throws, used to verify that PxlReflectionSupport.newClassInstance
 * propagates the constructor's own exception (the InvocationTargetException path).
 * Deliberately not a Lombok fixture: the throwing no-arg constructor is the whole point.
 */
public class ThrowingCtorRow {

    @PxlColumn(name = "Name")
    private String name;

    public ThrowingCtorRow() {
        throw new IllegalStateException("ctor-boom");
    }

    public String getName() {
        return name;
    }

    public void setName(final String name) {
        this.name = name;
    }

}
