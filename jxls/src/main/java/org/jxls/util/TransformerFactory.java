package org.jxls.util;

import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

import org.jxls.common.JxlsException;
import org.jxls.transform.Transformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Creates {@link Transformer} instances in runtime
 * @author Leonid Vysochyn
 */
public class TransformerFactory {
    public static final String POI_CLASS_NAME = "org.jxls.transform.poi.PoiTransformer";
    public static final String INIT_METHOD = "createTransformer";
    public static final String TRANSFORMER_SYSTEM_PROPERTY = "jxlstransformer";
    public static final String POI_TRANSFORMER = "poi";
    private static Logger logger = LoggerFactory.getLogger(TransformerFactory.class);

    /**
     * Creates a transformer initialized for reading a template from an {@link InputStream} and writing output to {@link OutputStream}
     * By default it creates a transformer from jxls-poi module
     * To create a different transformer set `jxlstransformer` system java property to the class name of the transformer
     * The transformer should have a public static `createTransformer(InputStream is, OutputStream os)` method
     * which will be invoked to create an instance of the transformer
     * @param inputStream - an input stream to read an Excel template
     * @param outputStream - an output stream to write the processed Excel output
     * @return {@link Transformer}
     */
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
            throw new JxlsException("The specified public method " + INIT_METHOD + " does not exist in " + transformer.getName());
        } catch (InvocationTargetException e) {
            if (e.getCause() instanceof CannotOpenWorkbookException) {
                throw (CannotOpenWorkbookException) e.getCause();
            }
            throw new JxlsException("Method " + INIT_METHOD + " of " + transformer.getName() + " class thrown an exception:\n" + e.getMessage(), e);
        } catch (IllegalAccessException e) {
            throw new JxlsException("Method " + INIT_METHOD + " of " + transformer.getName() + " is inaccessible", e);
        } catch (RuntimeException e) {
            throw new JxlsException("Failed to execute method " + INIT_METHOD + " of " + transformer.getName() + "\n" + e.getMessage(), e);
        }
    }

    /**
     * @deprecated Use {@link #getTransformerClassName()} method instead
     * @return a name for the {@link Transformer} which is loaded by this factory
     */
    public static String getTransformerName() {
        Class<?> transformerClass = getTransformerClass();
        if (transformerClass == null) {
            return null;
        }
        if (POI_CLASS_NAME.equalsIgnoreCase(transformerClass.getName())) {
            return POI_TRANSFORMER;
        }else{
            return transformerClass.getName();
        }
    }

    /**
     * @return the transformer class to be loaded by this factory
     */
    public static String getTransformerClassName() {
        String transformerClassName = System.getProperty(TRANSFORMER_SYSTEM_PROPERTY, POI_CLASS_NAME);
        // for backwards compatibility we also allow a short name for a poi transformer
        if (transformerClassName.equalsIgnoreCase(POI_TRANSFORMER)) {
            transformerClassName = POI_CLASS_NAME;
        }
        return transformerClassName;
    }

    private static Class<?> getTransformerClass() {
        String transformerClassName = getTransformerClassName();
        return loadTransformerByClass(transformerClassName);
    }

    private static Class<?> loadTransformerByClass(String transformerClassName) {
        try {
            logger.info("Loading transformer by class {}", transformerClassName);
            return Class.forName(transformerClassName);
        } catch (Exception e) {
            logger.error("Failed to load transformer class", e);
            throw new JxlsException("Failed to load transformer class " + transformerClassName, e);
        }
    }

}
