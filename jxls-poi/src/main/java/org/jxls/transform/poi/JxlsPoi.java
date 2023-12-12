package org.jxls.transform.poi;

import java.io.InputStream;
import java.util.Map;

import org.jxls.builder.JxlsOutput;
import org.jxls.builder.JxlsStreaming;

public class JxlsPoi {

    private JxlsPoi() {
    }
    
    public static JxlsOutput fill(InputStream template, JxlsStreaming streaming, Map<String, Object> data) {
        return JxlsPoiTemplateFillerBuilder.newInstance()
                .withStreaming(streaming)
                .withTemplate(template)
                .build()
                .fill(data);
    }
}
