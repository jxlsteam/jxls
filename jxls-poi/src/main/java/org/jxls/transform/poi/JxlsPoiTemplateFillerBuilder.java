package org.jxls.transform.poi;

import org.jxls.builder.JxlsTemplateFillerBuilder;
import org.jxls.common.PoiExceptionThrower;

public class JxlsPoiTemplateFillerBuilder extends JxlsTemplateFillerBuilder<JxlsPoiTemplateFillerBuilder> {

    public static JxlsPoiTemplateFillerBuilder newInstance() {
        return new JxlsPoiTemplateFillerBuilder().withTransformerFactory(new PoiTransformerFactory());
    }
    
    public JxlsPoiTemplateFillerBuilder withExceptionThrower() {
    	withExceptionHandler(new PoiExceptionThrower());
    	return this;
    }
}
