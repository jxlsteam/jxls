package org.jxls.util;

public interface ServiceFactory {
	
	public static final ServiceFactory DEFAULT = new ServiceFactoryDefaultImpl();
	
	<T> T createService(Class<T> interfaceClass, T defaultImpl);
}
