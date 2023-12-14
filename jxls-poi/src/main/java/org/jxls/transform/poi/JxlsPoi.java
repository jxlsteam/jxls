package org.jxls.transform.poi;

import java.io.File;
import java.io.InputStream;
import java.util.Map;

import org.jxls.builder.JxlsStreaming;

public class JxlsPoi {

    private JxlsPoi() {
    }
    
    public static void fill(InputStream template, JxlsStreaming streaming, Map<String, Object> data, File outputFile) {
        JxlsPoiTemplateFillerBuilder.newInstance()
                .withStreaming(streaming)
                .withTemplate(template)
                .buildAndFill(data, outputFile);
    }
}
