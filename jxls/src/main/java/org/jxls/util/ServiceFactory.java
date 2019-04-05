package org.jxls.util;

public interface ServiceFactory {
	
	ServiceFactory DEFAULT = new ServiceFactoryDefaultImpl();
	
	<T> T createService(Class<T> interfaceClass, T defaultImpl);
}
