package io.github.hclimkr.pxl.tcdata;

import lombok.*;

/**
 * Custom object type.
 * Import round-trips via a single-arg String constructor, export via the overridden toString.
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class Point {

    private int x;

    private int y;

    // String -> Point (import path)
    public Point(final String value) {
        final String[] parts = value.split(",");
        this.x = Integer.parseInt(parts[0].trim());
        this.y = Integer.parseInt(parts[1].trim());
    }

    // Point -> String (export path)
    @Override
    public String toString() {
        return this.x + "," + this.y;
    }

}
