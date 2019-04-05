package org.jxls.util;

import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

import org.jxls.transform.Transformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn
 * @since 11/6/13
 */
public class TransformerFactory {
    public static final String POI_CLASS_NAME = "org.jxls.transform.poi.PoiTransformer";
    /**
     * @deprecated Use POI_CLASS_NAME. JEXCEL_CLASS_NAME might be removed in a future release.
     */
    public static final String JEXCEL_CLASS_NAME = "org.jxls.transform.jexcel.JexcelTransformer";
    public static final String INIT_METHOD = "createTransformer";
    public static final String TRANSFORMER_SYSTEM_PROPERTY = "jxlstransformer";
    /**
     * @deprecated Use POI_TRANSFORMER. JEXCEL_TRANSFORMER might be removed in a future release.
     */
    public static final String JEXCEL_TRANSFORMER = "jexcel";
    public static final String POI_TRANSFORMER = "poi";
    private static Logger logger = LoggerFactory.getLogger(TransformerFactory.class);

    public static Transformer createTransformer(InputStream inputStream, OutputStream outputStream) {
        Class<?> transformer = getTransformerClass();
        if (transformer == null) {
            logger.error("Cannot load any Transformer class. Please make sure you have necessary libraries in CLASSPATH.");
            return null;
        }
        logger.debug("Transformer class is " + transformer.getName());
        try {
            Method initMethod = transformer.getMethod(INIT_METHOD, InputStream.class, OutputStream.class);
            return (Transformer) initMethod.invoke(null, inputStream, outputStream);
        } catch (NoSuchMethodException e) {
            logger.error("The specified public method " + INIT_METHOD + " does not exist in " + transformer.getName());
            return null;
        } catch (InvocationTargetException e) {
            logger.error("Method " + INIT_METHOD + " of " + transformer.getName() + " class thrown an Exception", e);
            return null;
        } catch (IllegalAccessException e) {
            logger.error("Method " + INIT_METHOD + " of " + transformer.getName() + " is inaccessible", e);
            return null;
        } catch (RuntimeException e) {
            logger.error("Failed to execute method " + INIT_METHOD + " of " + transformer.getName(), e);
            return null;
        }
    }

    public static String getTransformerName() {
        Class<?> transformerClass = getTransformerClass();
        if (transformerClass == null) {
            return null;
        }
        if (POI_CLASS_NAME.equalsIgnoreCase(transformerClass.getName())) {
            return POI_TRANSFORMER;
        }
        if (JEXCEL_CLASS_NAME.equalsIgnoreCase(transformerClass.getName())) {
            logger.warn("jxls-jexcel is deprecated");
            return JEXCEL_TRANSFORMER;
        }
        return transformerClass.getName();
    }

    private static Class<?> getTransformerClass() {
        String transformerName = System.getProperty(TRANSFORMER_SYSTEM_PROPERTY);
        Class<?> transformer = null;
        if (transformerName == null) {
            transformer = loadPoiTransformer();
            if (transformer == null) {
                transformer = loadJexcelTransformer();
            }
        } else {
            if (POI_TRANSFORMER.equalsIgnoreCase(transformerName)) {
                transformer = loadPoiTransformer();
            } else if (JEXCEL_TRANSFORMER.equalsIgnoreCase(transformerName)) {
                transformer = loadJexcelTransformer();
            }
        }
        return transformer;
    }

    private static Class<?> loadPoiTransformer() {
        try {
            return Class.forName(POI_CLASS_NAME);
        } catch (Exception e) {
            logger.warn("Cannot load POI transformer class", e);
            return null;
        }
    }

    private static Class<?> loadJexcelTransformer() {
        logger.warn("jxls-jexcel is deprecated");
        try {
            return Class.forName(JEXCEL_CLASS_NAME);
        } catch (Exception e) {
            logger.warn("Cannot load JExcel transformer class", e);
            return null;
        }
    }
}
