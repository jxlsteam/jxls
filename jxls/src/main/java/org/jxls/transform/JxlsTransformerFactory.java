package org.jxls.transform;

import java.io.InputStream;
import java.io.OutputStream;

import org.jxls.builder.JxlsStreaming;

public interface JxlsTransformerFactory {

    Transformer create(InputStream template, OutputStream outputStream, JxlsStreaming streaming);
}
