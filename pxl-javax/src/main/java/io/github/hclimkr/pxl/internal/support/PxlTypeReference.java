package io.github.hclimkr.pxl.internal.support;

import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnostic;
import io.github.hclimkr.pxl.internal.i18n.PxlI18nDiagnosticKeys;

import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;

/**
 * Super type token
 */
@Deprecated
public abstract class PxlTypeReference<T> {

    private final Type type;

    /**
     * Captures the type argument {@code T} from an anonymous subclass's generic superclass.
     *
     * @throws IllegalArgumentException if instantiated without an actual type argument (raw superclass)
     */
    protected PxlTypeReference() {

        final Type superClass = getClass().getGenericSuperclass();
        if (superClass instanceof Class<?>) {
            throw new IllegalArgumentException(PxlI18nDiagnostic.get(PxlI18nDiagnosticKeys.SUPPORT_TYPE_REFERENCE_MISSING));
        }

        type = ((ParameterizedType) superClass).getActualTypeArguments()[0];
    }

    /**
     * Creates a reference wrapping an explicitly supplied type.
     *
     * @param type the captured type
     */
    private PxlTypeReference(Type type) {

        this.type = type;
    }

    /**
     * Returns the captured type.
     *
     * @return the captured type
     */
    public Type getType() {

        return type;
    }

}
