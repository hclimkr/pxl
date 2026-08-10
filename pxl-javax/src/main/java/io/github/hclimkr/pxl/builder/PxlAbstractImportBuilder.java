package io.github.hclimkr.pxl.builder;

import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;
import org.apache.commons.io.FilenameUtils;

import javax.validation.Validator;
import java.io.File;
import java.text.Normalizer;

/**
 * Common base for import builders.
 *
 * <p>Holds common state independent of the source (Excel/CSV) - {@code validator}, {@code workbookName},
 * {@code option} - and the shared rule for deriving a name from a source file ({@link #getNormalizedFileBaseName(File)}).
 * The parse-target configuration ({@code workbook(...)}/{@code sheet(...)}), which returns a source-terminal step
 * ({@code Pxl*ImportBuilder.Source}), and the setters {@code workbookName(...)}/{@code override(...)} - which must return
 * the self type for chaining - are implemented by each source's subclass builder, which sets these fields. The source
 * step carries its own copy of {@code workbookName}/{@code option} and re-declares the same setters, so they may be
 * chained after the parse-target configuration as well (the value set last wins).</p>
 *
 * <p>Package-private: not part of the public API. It exposes no public members; the public chaining methods live on
 * the public concrete subclasses.</p>
 */
abstract class PxlAbstractImportBuilder {

    /**
     * The shared bean-validation validator, or {@code null} when bean validation is disabled.
     */
    protected final Validator validator;

    /**
     * The name bound to the workbook object's {@code @PxlWorkbookName} field in the workbook form; {@code null} until
     * set by the subclass setter, in which case an Excel import from a file falls back to the source file name.
     */
    protected String workbookName;

    /**
     * The import workbook option overriding annotation values; {@code null} when unset.
     */
    protected PxlImportWorkbookOption option;

    /**
     * Stores the shared bean-validation validator for the subclass builder.
     *
     * @param validator the bean-validation validator, or {@code null} when bean validation is disabled
     */
    protected PxlAbstractImportBuilder(final Validator validator) {

        this.validator = validator;
    }

    /**
     * Derives a name from a source file: its file name without the extension, normalized to Unicode NFC and trimmed.
     *
     * <p>Both import sources name things after files - Excel fills the {@code @PxlWorkbookName} field from the source
     * file, CSV matches each file against the {@code @PxlSheet} names - so both derive the name the same way.</p>
     *
     * <p>The NFC step matters on macOS, whose file systems hand back <em>decomposed</em> (NFD) file names: a Korean
     * name would otherwise not equal its composed counterpart written in an annotation, and the match would fail even
     * though both read identically. Windows and Linux hand back composed names already, so normalizing is a no-op there.</p>
     *
     * @param file the source file
     * @return the file's base name in NFC, trimmed; empty when the file name is nothing but an extension (e.g. {@code ".xlsx"})
     */
    protected static String getNormalizedFileBaseName(final File file) {

        return Normalizer.normalize(FilenameUtils.getBaseName(file.getName()), Normalizer.Form.NFC).trim();
    }

}
