package org.jxls.util;

import java.util.Iterator;
import java.util.ServiceLoader;

import org.jxls.template.SimpleExporter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ServiceFactoryDefaultImpl implements ServiceFactory {

	private final static Logger LOG = LoggerFactory.getLogger(SimpleExporter.class);
	
	@Override
	public <T> T createService(Class<T> interfaceClass, T defaultImpl) {
		
		final Iterator<T> iterator = ServiceLoader.load(interfaceClass).iterator();
		final T ret;
		if( iterator.hasNext() ){
			ret = iterator.next();
		}else{
			ret = defaultImpl;
		}
		LOG.debug("SPI {} => {}",interfaceClass.getName(),ret );
		LOG.info("you may change the SPI on file: META-INF/services/{}",interfaceClass.getName());
		return ret;
	}

}
