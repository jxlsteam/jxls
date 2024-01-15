package org.jxls.common;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Map;

import org.apache.commons.beanutils.PropertyUtils;
import org.jxls.logging.JxlsLogger;

public class ObjectPropertyAccess {

    private ObjectPropertyAccess() {
    }
    
    /**
     * Dynamically sets an object property via reflection
     * @param obj -
     * @param propertyName -
     * @param propertyValue -
     */
    public static void setObjectProperty(Object obj, String propertyName, String propertyValue, JxlsLogger logger) {
        try {
            setObjectProperty(obj, propertyName, propertyValue);
        } catch (Exception e) {
            logger.handleSetObjectPropertyException(e, obj, propertyName, propertyValue);
        }
    }

    /**
     * Dynamically sets an object property via reflection
     * @param obj -
     * @param propertyName -
     * @param propertyValue -
     * @throws NoSuchMethodException -
     * @throws InvocationTargetException -
     * @throws IllegalAccessException -
     */
    public static void setObjectProperty(Object obj, String propertyName, String propertyValue)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        String name = "set" + Character.toUpperCase(propertyName.charAt(0)) + propertyName.substring(1);
        Method method = obj.getClass().getMethod(name, new Class[] { String.class });
        method.invoke(obj, propertyValue);
    }

    /**
     * Gets value of the passed object by the given property name.
     * @param obj Map, DynaBean or Java bean
     * @param propertyName -
     * @return value (can be null)
     */
    public static Object getObjectProperty(Object obj, String propertyName, JxlsLogger logger) {
        try {
            return getObjectProperty(obj, propertyName);
        } catch (Exception e) {
            logger.handleGetObjectPropertyException(e, obj, propertyName);
            return null;
        }
    }

    /**
     * Gets value of the passed object by the given property name.
     * @param obj Map, DynaBean or Java bean
     * @param propertyName -
     * @return value (can be null)
     * @throws NoSuchMethodException -
     * @throws InvocationTargetException -
     * @throws IllegalAccessException -
     */
    public static Object getObjectProperty(Object obj, String propertyName) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        if (obj instanceof Map<?,?> map) { // Map access
            return map.get(propertyName);
        } else { // DynaBean or Java bean access
            return PropertyUtils.getProperty(obj, propertyName);
        }
    }
}
