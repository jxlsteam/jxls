package org.jxls.util;

public class JxlsConfigProviderImpl implements JxlsConfigProvider {

    @Override
    public String getProperty(String key, String defaultValue) {
        return System.getProperty(key, defaultValue);
    }
}
