package org.jxls.util;

/**
 * {@link JxlsConfigProvider} based on java system properties
 */
public class JxlsConfigProviderImpl implements JxlsConfigProvider {

    @Override
    public String getProperty(String key, String defaultValue) {
        return System.getProperty(key, defaultValue);
    }
}
