package org.jxls.transform.poi;

import org.jxls.builder.JxlsTemplateFillerBuilder;

public class JxlsPoiTemplateFillerBuilder extends JxlsTemplateFillerBuilder<JxlsPoiTemplateFillerBuilder> {

    public static JxlsPoiTemplateFillerBuilder newInstance() {
        return new JxlsPoiTemplateFillerBuilder().withTransformerFactory(new PoiTransformerFactory());
    }
}
