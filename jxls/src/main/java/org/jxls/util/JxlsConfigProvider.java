package org.jxls.util;

/**
 * Provides Jxls configuration properties
 */
public interface JxlsConfigProvider {
	
	String getProperty(String key, String defaultValue);
}
