package io.github.hclimkr.pxl.exception;

/**
 * Thrown when the workbook/CSV shape does not match what the binding expects.
 *
 * <p>Covers a target that cannot be located or is ambiguous (no sheet matching the candidate names, a missing
 * column, a duplicated sheet/column name, a CSV name matching several sheets), a structural limit that is
 * exceeded (maximum number of sheets/rows/columns), a missing header row/column, nothing to write on export
 * (no data, a {@code null} row collection or row class), inputs whose counts do not line up (sheet names vs.
 * row classes/collections, CSV names vs. streams), and row/column index settings that cannot hold together
 * (a negative index, a last index below the first, a first data row not after the header row).</p>
 */
public final class PxlDataException extends PxlException {

    /**
     * Creates an exception with no detail message.
     */
    public PxlDataException() {

        super();
    }

    /**
     * Creates an exception with the given detail message.
     *
     * @param message the detail message
     */
    public PxlDataException(final String message) {

        super(message);
    }

    /**
     * Creates an exception with the given detail message and cause.
     *
     * @param message the detail message
     * @param cause   the underlying cause
     */
    public PxlDataException(final String message, final Throwable cause) {

        super(message, cause);
    }

    /**
     * Creates an exception with the given cause.
     *
     * @param cause the underlying cause
     */
    public PxlDataException(final Throwable cause) {

        super(cause);
    }

}
