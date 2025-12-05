package org.jxls.transform.poi;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

import org.jxls.builder.JxlsStreaming;
import org.jxls.common.PoiExceptionLogger;

public class JxlsPoi {

    private JxlsPoi() {
    }
    
    public static void fill(InputStream template, JxlsStreaming streaming, Map<String, Object> data, File outputFile) {
        JxlsPoiTemplateFillerBuilder.newInstance()
                .withStreaming(streaming)
                .withTemplate(template)
                .buildAndFill(data, outputFile);
    }

    public static void fill(InputStream template, JxlsStreaming streaming, Map<String, Object> data, OutputStream out) {
        JxlsPoiTemplateFillerBuilder.newInstance()
                .withStreaming(streaming)
                .withTemplate(template)
                .buildAndFill(data, () -> out);
    }
    
    /**
     * Use this method if you just need PoiTransformer functionality for a template file.
     * @param template -
     * @return PoiTransformer
     */
    public static PoiTransformer createSimple(InputStream template) {
        return (PoiTransformer) new PoiTransformerFactory().create(template, new ByteArrayOutputStream(),
                    JxlsStreaming.STREAMING_OFF, List.of(), new PoiExceptionLogger());
    }
}
