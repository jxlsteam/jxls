package org.jxls.expression;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.jexl3.JexlEngine;
import org.apache.commons.jexl3.JexlException;
import org.apache.commons.jexl3.JexlOperator;
import org.apache.commons.jexl3.introspection.JexlPropertyGet;
import org.apache.commons.jexl3.introspection.JexlPropertySet;
import org.apache.commons.jexl3.introspection.JexlUberspect;

/**
 * A {@link JexlUberspect.ResolverStrategy} that extends the standard POJO resolution
 * with Java Record support.
 * 
 * <p>Standard JEXL introspection resolves {@code obj.prop} by looking for {@code getProp()} or
 * {@code isProp()} methods (JavaBeans convention). Java Records use accessor methods named
 * {@code prop()} without any prefix, so the standard resolution fails.</p>
 * 
 * <p>This strategy appends a custom {@link RecordPropertyResolver} after the standard POJO
 * resolvers. When all standard resolvers fail and the object is a {@link java.lang.Record},
 * it attempts to find a no-argument method with exactly the property name.</p>
 */
public class RecordAwareStrategy implements JexlUberspect.ResolverStrategy {

    private static final RecordPropertyResolver RECORD_RESOLVER = new RecordPropertyResolver();

    @Override
    public List<JexlUberspect.PropertyResolver> apply(final JexlOperator operator, final Object obj) {
        List<JexlUberspect.PropertyResolver> resolvers = new ArrayList<>(JexlUberspect.POJO);
        resolvers.add(RECORD_RESOLVER);
        return resolvers;
    }

    /**
     * A {@link JexlUberspect.PropertyResolver} that handles Java Record component accessors.
     * <p>When the object is a {@link Record}, it looks for a no-argument method
     * with the exact property name (e.g. {@code name()} for property {@code name}).
     * For property sets, it returns {@code null} since records are immutable.</p>
     */
    private static final class RecordPropertyResolver implements JexlUberspect.PropertyResolver {

        @Override
        public JexlPropertyGet getPropertyGet(final JexlUberspect uber, final Object obj,
                                              final Object identifier) {
            if (obj instanceof Record) {
                String propertyName = identifier.toString();
                try {
                    Method method = obj.getClass().getMethod(propertyName);
                    if (method.getParameterCount() == 0 && method.getReturnType() != void.class) {
                        return new RecordPropertyGet(obj.getClass(), method, propertyName);
                    }
                } catch (final NoSuchMethodException e) {
                }
            }
            return null;
        }

        @Override
        public JexlPropertySet getPropertySet(final JexlUberspect uber, final Object obj,
                                              final Object identifier, final Object arg) {
            if (obj instanceof Record) {
                return null;
            }
            return null;
        }
    }

    /**
     * A {@link JexlPropertyGet} that invokes a record component accessor method.
     */
    private static final class RecordPropertyGet implements JexlPropertyGet {

        private final Class<?> objectClass;
        private final Method method;
        private final String propertyName;

        RecordPropertyGet(final Class<?> objectClass, final Method method, final String propertyName) {
            this.objectClass = objectClass;
            this.method = method;
            this.propertyName = propertyName;
        }

        @Override
        public Object invoke(final Object obj) throws Exception {
            return method.invoke(obj);
        }

        @Override
        public boolean isCacheable() {
            return true;
        }

        @Override
        public boolean tryFailed(final Object rval) {
            return rval == JexlEngine.TRY_FAILED;
        }

        @Override
        public Object tryInvoke(final Object obj, final Object key)
                throws JexlException.TryFailed {
            if (obj != null
                    && objectClass.equals(obj.getClass())
                    && propertyName.equals(key)) {
                try {
                    return method.invoke(obj);
                } catch (final IllegalAccessException | IllegalArgumentException e) {
                    return JexlEngine.TRY_FAILED;
                } catch (final InvocationTargetException e) {
                    throw JexlException.tryFailed(e);
                }
            }
            return JexlEngine.TRY_FAILED;
        }

        @Override
        public String toString() {
            return "RecordPropertyGet[" + objectClass.getSimpleName() + "." + propertyName + "()]";
        }
    }
}
