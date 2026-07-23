package io.github.hclimkr.pxl.tcdata;

import org.junit.jupiter.api.TestInfo;

import java.io.File;
import java.lang.reflect.Method;

/**
 * Common path for storing test export artifacts.
 * <p>
 * All file-based export tests leave actual files under {@code target/test-outputs/}.
 * (They persist after the tests so a human can open them, and {@code mvn clean} deletes the whole
 * {@code target/} directory, so these artifacts are removed with it.)
 * <p>
 * export→import round-trip tests match the artifact file name to the <b>test method name</b>.
 * Passing the {@link TestInfo} injected in {@link org.junit.jupiter.api.BeforeEach}
 * returns a {@code <methodName>.xlsx} (or the given extension) file handle.
 */
public final class TestPaths {

    public static final String EXPORT_DIR = "target/test-outputs";

    private TestPaths() {
    }

    /**
     * Returns a file handle under the export folder (creating the folder if it does not exist).
     */
    public static File exportFile(final String name) {
        final File dir = new File(EXPORT_DIR);
        dir.mkdirs();
        return new File(dir, name);
    }

    /**
     * Returns a {@code <methodName>.xlsx} file handle using the current test method name.
     */
    public static File exportFile(final TestInfo testInfo) {
        return exportFile(testInfo, ".xlsx");
    }

    /**
     * Returns a file handle using the current test method name + the given extension (e.g. {@code ".xls"}).
     */
    public static File exportFile(final TestInfo testInfo, final String extension) {
        return exportFile(methodName(testInfo) + extension);
    }

    private static String methodName(final TestInfo testInfo) {
        return testInfo.getTestMethod()
                .map(Method::getName)
                .orElseThrow(() -> new IllegalStateException("could not determine the test method name"));
    }

}
