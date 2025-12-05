package org.jxls.transform;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.jxls.builder.JxlsStreaming;
import org.jxls.logging.JxlsLogger;

public interface JxlsTransformerFactory {

    /**
     * @param template -
     * @param outputStream -
     * @param streaming -
     * @param templatePreprocessors should be called after template is opened and before Transformer is instantiated
     * @param logger -
     * @return new Transformer instance
     */
    Transformer create(InputStream template, OutputStream outputStream, JxlsStreaming streaming,
            List<TemplateProcessor> templatePreprocessors, JxlsLogger logger);
}
