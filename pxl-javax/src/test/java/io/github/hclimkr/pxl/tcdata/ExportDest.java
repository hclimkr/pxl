package io.github.hclimkr.pxl.tcdata;

/**
 * The terminal destinations an export builder can be driven to, swept by the parameterized export tests.
 * <p>
 * {@code FILE} and {@code STREAM} exist on every export builder; {@code WORKBOOK} only on the Excel ones
 * ({@code toWorkbook()} has no CSV counterpart), so CSV tests narrow the sweep with
 * {@code @EnumSource(value = ExportDest.class, names = {"FILE", "STREAM"})}.
 *
 * @see TestExports
 */
public enum ExportDest {

    /**
     * {@code toFile(File)} - the result is read back from the written file.
     */
    FILE,

    /**
     * {@code toStream(OutputStream)} - the result is read back from the bytes written to the stream.
     */
    STREAM,

    /**
     * {@code toWorkbook()} - Excel only; hands the caller a live POI workbook instead of writing anywhere.
     */
    WORKBOOK

}
