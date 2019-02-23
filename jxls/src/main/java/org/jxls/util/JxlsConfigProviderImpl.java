package org.jxls.util;

public class JxlsConfigProviderImpl implements JxlsConfigProvider {

	public String getProperty(String key, String defaultValue) {
		return System.getProperty(key, defaultValue);
	}

}
