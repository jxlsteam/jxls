package org.jxls.builder;

import java.io.ByteArrayOutputStream;
import java.io.OutputStream;

public class JxlsOutputByteArray implements JxlsOutput {

    private final ByteArrayOutputStream byteArrayOutputStream;

    public JxlsOutputByteArray(ByteArrayOutputStream byteArrayOutputStream) {
        this.byteArrayOutputStream = byteArrayOutputStream;
    }

    @Override
    public OutputStream getOutputStream() {
        return byteArrayOutputStream;
    }

}