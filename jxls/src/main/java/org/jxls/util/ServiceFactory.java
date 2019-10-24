package org.jxls.util;

/**
 * Service factory to load different SPI
 */
public interface ServiceFactory {
	
	ServiceFactory DEFAULT = new ServiceFactoryDefaultImpl();
	
	<T> T createService(Class<T> interfaceClass, T defaultImpl);
}
