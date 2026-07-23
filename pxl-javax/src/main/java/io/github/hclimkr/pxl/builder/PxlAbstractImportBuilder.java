package io.github.hclimkr.pxl.builder;

import io.github.hclimkr.pxl.option.PxlImportWorkbookOption;

import javax.validation.Validator;

/**
 * Common base for import builders.
 *
 * <p>Holds common state independent of the source (Excel/CSV): {@code validator}, {@code workbookName}, {@code option}.
 * The parse-target configuration ({@code workbook(...)}/{@code sheet(...)}), which returns a source-terminal step
 * ({@code Pxl*ImportBuilder.Source}), and the setters {@code workbookName(...)}/{@code override(...)} — which must return
 * the self type for chaining — are implemented by each source's subclass builder, which sets these fields.</p>
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
     * The workbook name used to resolve the target sheet(s); {@code null} until set by the subclass setter.
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

}
