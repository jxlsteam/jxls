package org.jxls.util;

import org.jxls.transform.Transformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 * @author Leonid Vysochyn
 *         Date: 11/6/13
 */
public class TransformerFactory {
    public static final String POI_CLASS_NAME = "org.jxls.transform.poi.PoiTransformer";
    public static final String JEXCEL_CLASS_NAME = "org.jxls.transform.jexcel.JexcelTransformer";
    public static final String INIT_METHOD = "createTransformer";
    public static final String TRANSFORMER_SYSTEM_PROPERTY = "jxlstransformer";
    public static final String JEXCEL_TRANSFORMER = "jexcel";
    public static final String POI_TRANSFORMER = "poi";

    private static Logger logger = LoggerFactory.getLogger(TransformerFactory.class);

    public static Transformer createTransformer(InputStream inputStream, OutputStream outputStream) {
        Class transformer = getTransformerClass();
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
        } catch (RuntimeException e){
            logger.error("Failed to execute method " + INIT_METHOD + " of " + transformer.getName(), e);
            return null;
        }
    }

    public static String getTransformerName(){
        Class transformerClass = getTransformerClass();
        if( transformerClass == null ){
            return null;
        }
        if( POI_CLASS_NAME.equalsIgnoreCase( transformerClass.getName() ) ){
            return POI_TRANSFORMER;
        }
        if( JEXCEL_CLASS_NAME.equalsIgnoreCase( transformerClass.getName() ) ){
            return JEXCEL_TRANSFORMER;
        }
        return transformerClass.getName();
    }

    private static Class getTransformerClass() {
        String transformerName = System.getProperty(TRANSFORMER_SYSTEM_PROPERTY);
        Class transformer = null;
        if( transformerName == null ){
            transformer = loadPoiTransformer();
            if (transformer == null) {
                transformer = loadJexcelTransformer();
            }
        }else{
            if( POI_TRANSFORMER.equalsIgnoreCase(transformerName ) ){
                transformer = loadPoiTransformer();
            }else if( JEXCEL_TRANSFORMER.equalsIgnoreCase(transformerName)){
                transformer = loadJexcelTransformer();
            }
        }
        return transformer;
    }

    private static Class loadPoiTransformer() {
        try {
            return Class.forName(POI_CLASS_NAME);
        } catch (Exception e) {
            logger.info("Cannot load POI transformer class");
            return null;
        }
    }

    private static Class loadJexcelTransformer() {
        try {
            return Class.forName(JEXCEL_CLASS_NAME);
        } catch (Exception e) {
            logger.info("Cannot load JExcel transformer class");
            return null;
        }
    }

}
