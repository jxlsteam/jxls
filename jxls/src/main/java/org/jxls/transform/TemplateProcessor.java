package org.jxls.transform;

import org.jxls.builder.JxlsStreaming;
import org.jxls.logging.JxlsLogger;

public interface TemplateProcessor {

    /**
     * @param template type is implementation specific
     * @param streaming -
     * @param logger -
     */
    void process(Object template, JxlsStreaming streaming, JxlsLogger logger);
}
