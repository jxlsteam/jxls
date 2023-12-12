package org.jxls.builder;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class JxlsOutput {
    private final ByteArrayOutputStream outputStream;
    
    protected JxlsOutput(OutputStream outputStream) {
        this.outputStream = (ByteArrayOutputStream) outputStream;
    }
    
    public byte[] getByteArray() {
        return outputStream.toByteArray();
    }
    
    public void write(File file) {
        try (FileOutputStream fos = new FileOutputStream(file)) {
            fos.write(getByteArray());
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
}
