package org.jxls.builder;

import java.io.IOException;

public class JxlsTemplateFillException extends RuntimeException {

    public JxlsTemplateFillException(IOException ex) {
        super(ex);
    }
}
