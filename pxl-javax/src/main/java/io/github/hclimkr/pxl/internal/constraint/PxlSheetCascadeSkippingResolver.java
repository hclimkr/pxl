package io.github.hclimkr.pxl.internal.constraint;

import io.github.hclimkr.pxl.annotation.PxlSheet;
import io.github.hclimkr.pxl.internal.support.PxlReflectionSupport;

import javax.validation.Path;
import javax.validation.TraversableResolver;
import java.lang.annotation.ElementType;
import java.lang.reflect.Field;
import java.util.HashSet;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Traversable resolver that stops bean validation from cascading into a workbook object's {@link PxlSheet} fields,
 * delegating every other decision to the provider's default resolver.
 *
 * <p>Rationale: the binder validates row objects itself, one by one, tagging each violation with the sheet name
 * and - on import - the row index. When a {@code @PxlSheet} field is also marked {@code @Valid}, validating the
 * workbook object cascades into those very same rows a second time. The outcome is identical (validation is
 * idempotent), so the second traversal only costs time, and on export it costs diagnostics too: the workbook pass
 * runs first and carries no location, so it reports the violation untagged.</p>
 *
 * <p>Blocking that cascade leaves the binder's own per-row pass as the single place sheet rows are validated, which
 * has a useful consequence: rows are validated exactly where they are processed, and nowhere else. That pass sits
 * behind the binder's own enabled check, so a sheet the binder skips - {@code exportEnabled = false} on export,
 * {@code importEnabled = false} on import - has its rows left unvalidated whether or not the field carries
 * {@code @Valid}. Validation follows what is actually written or read.</p>
 *
 * <p>Everything else is untouched: the workbook object's own constraints, the constraints declared on the sheet
 * collection itself ({@code @NotEmpty}, {@code @Size}, ...), and any {@code @Valid} on a field that is not a
 * sheet.</p>
 *
 * <p>Only a sheet field of the <em>root</em> object is skipped, because only there is the field a sheet PXL is
 * about to bind. The same declaration sitting on an object reached below the root belongs to a graph the binder
 * knows nothing about, so the {@code @Valid} on it means what it ordinarily means and is left alone.</p>
 *
 * @see TraversableResolver#isCascadable(Object, Path.Node, Class, Path, ElementType)
 */
public final class PxlSheetCascadeSkippingResolver implements TraversableResolver {

    /**
     * Stand-in used when no default resolver is available, matching the Bean Validation default of traversing
     * everything. The specification has {@code Configuration} return a default resolver, so this is a guard
     * against a provider that does not, not an expected path.
     */
    private static final TraversableResolver TRAVERSE_ALL = new TraverseAllResolver();

    /**
     * The resolver every decision but the sheet-field cascade is handed to; never {@code null}, falling back to
     * {@code TRAVERSE_ALL} when the caller supplies none.
     */
    private final TraversableResolver delegate;

    /**
     * Per-class cache of the {@link PxlSheet} field names, so the reflective scan runs once per workbook class
     * rather than once per cascade decision.
     */
    private final Map<Class<?>, Set<String>> sheetFieldNamesByClass = new ConcurrentHashMap<>();

    /**
     * Creates a resolver wrapping the given default resolver.
     *
     * @param delegate the provider's default traversable resolver, which every decision but the sheet-field
     *                 cascade is delegated to; the traverse-everything stand-in is used when {@code null}
     */
    public PxlSheetCascadeSkippingResolver(final TraversableResolver delegate) {

        this.delegate = Objects.nonNull(delegate) ? delegate : TRAVERSE_ALL;
    }

    /**
     * Always delegates: whether a property is reachable is none of this resolver's business.
     *
     * @param traversableObject       the object hosting the property
     * @param traversableProperty     the property being traversed
     * @param rootBeanType            the type of the validated root object
     * @param pathToTraversableObject the path from the root object to the hosting object
     * @param elementType             {@code FIELD} or {@code METHOD}
     * @return whatever the delegate answers
     */
    @Override
    public boolean isReachable(final Object traversableObject,
                               final Path.Node traversableProperty,
                               final Class<?> rootBeanType,
                               final Path pathToTraversableObject,
                               final ElementType elementType) {

        return delegate.isReachable(traversableObject, traversableProperty, rootBeanType, pathToTraversableObject, elementType);
    }

    /**
     * Answers {@code false} - refusing the cascade - for a {@link PxlSheet} field of the validated root, and
     * delegates every other case.
     *
     * <p>The provider calls this only for a property already marked {@code @Valid}, which is what makes it the
     * cascade switch. Refusing here leaves the rows to the binder's own per-row validation; see the class
     * documentation for why that is the better of the two passes.</p>
     *
     * @param traversableObject       the object hosting the property, or {@code null} when the provider has no
     *                                instance at hand (as with {@code validateValue})
     * @param traversableProperty     the property being considered for cascading
     * @param rootBeanType            the type of the validated root object
     * @param pathToTraversableObject the path from the root object to the hosting object; a path naming no
     *                                property means the hosting object is the root
     * @param elementType             {@code FIELD} or {@code METHOD}, depending on where {@code @Valid} sits
     * @return {@code false} for a sheet field of the root, otherwise whatever the delegate answers
     */
    @Override
    public boolean isCascadable(final Object traversableObject,
                                final Path.Node traversableProperty,
                                final Class<?> rootBeanType,
                                final Path pathToTraversableObject,
                                final ElementType elementType) {

        if (isRootLevel(pathToTraversableObject) && isSheetField(traversableObject, rootBeanType, traversableProperty)) {
            return false;
        }

        return delegate.isCascadable(traversableObject, traversableProperty, rootBeanType, pathToTraversableObject, elementType);
    }

    /**
     * Tells whether the object hosting the property is the validated root itself, i.e. the path leading to it
     * names no property.
     *
     * @param pathToTraversableObject the path from the root object to the object hosting the property; {@code null}
     *                                is read as the root, since there is then no step away from it
     * @return {@code true} when the hosting object is the root
     */
    private static boolean isRootLevel(final Path pathToTraversableObject) {

        if (Objects.isNull(pathToTraversableObject)) {
            return true;
        }

        for (final Path.Node node : pathToTraversableObject) {
            if (Objects.nonNull(node.getName())) {
                return false;
            }
        }

        return true;
    }

    /**
     * Tells whether the given property is a {@link PxlSheet} field of the object hosting it.
     *
     * @param traversableObject   the object hosting the property, or {@code null} when the provider has no instance
     * @param rootBeanType        the type of the validated root object, used when no hosting instance is available
     * @param traversableProperty the property being considered for cascading
     * @return {@code true} when the property is a sheet field
     */
    private boolean isSheetField(final Object traversableObject,
                                 final Class<?> rootBeanType,
                                 final Path.Node traversableProperty) {

        final Class<?> hostType = Objects.nonNull(traversableObject) ? traversableObject.getClass() : rootBeanType;
        if (Objects.isNull(hostType) || Objects.isNull(traversableProperty)) {
            return false;
        }

        final String propertyName = traversableProperty.getName();
        if (Objects.isNull(propertyName)) {
            return false;
        }

        return sheetFieldNames(hostType).contains(propertyName);
    }

    /**
     * Returns the names of the {@link PxlSheet} fields declared by the given class or any of its superclasses.
     *
     * <p>The result is cached per class, so the reflective scan runs once per workbook class for the lifetime of
     * this resolver.</p>
     *
     * @param hostType the class to scan
     * @return the sheet field names, empty when the class declares none
     */
    private Set<String> sheetFieldNames(final Class<?> hostType) {

        return sheetFieldNamesByClass.computeIfAbsent(hostType, type -> {
            final Set<String> names = new HashSet<>();

            for (final Field field : PxlReflectionSupport.getAllFields(type)) {
                if (field.isAnnotationPresent(PxlSheet.class)) {
                    names.add(field.getName());
                }
            }

            return names;
        });
    }

    /**
     * Traverse-everything resolver matching the Bean Validation default, used when the provider hands out no
     * default resolver to delegate to.
     */
    private static final class TraverseAllResolver implements TraversableResolver {

        @Override
        public boolean isReachable(final Object traversableObject,
                                   final Path.Node traversableProperty,
                                   final Class<?> rootBeanType,
                                   final Path pathToTraversableObject,
                                   final ElementType elementType) {

            return true;
        }

        @Override
        public boolean isCascadable(final Object traversableObject,
                                    final Path.Node traversableProperty,
                                    final Class<?> rootBeanType,
                                    final Path pathToTraversableObject,
                                    final ElementType elementType) {

            return true;
        }

    }

}
