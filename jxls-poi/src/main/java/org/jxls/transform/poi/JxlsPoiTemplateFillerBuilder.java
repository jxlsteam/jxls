package org.jxls.transform.poi;

import org.jxls.builder.JxlsTemplateFillerBuilder;
import org.jxls.common.PoiExceptionLogger;
import org.jxls.common.PoiExceptionThrower;

public class JxlsPoiTemplateFillerBuilder extends JxlsTemplateFillerBuilder<JxlsPoiTemplateFillerBuilder> {

    public static JxlsPoiTemplateFillerBuilder newInstance() {
        return new JxlsPoiTemplateFillerBuilder().withLogger(new PoiExceptionLogger()).withTransformerFactory(new PoiTransformerFactory());
    }
    
    public JxlsPoiTemplateFillerBuilder withExceptionThrower() {
    	withLogger(new PoiExceptionThrower());
    	return this;
    }
}
