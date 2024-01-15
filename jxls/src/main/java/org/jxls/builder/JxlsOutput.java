package org.jxls.builder;

import java.io.IOException;
import java.io.OutputStream;

/**
 * An implementation is used to deliver the OutputStream for writing the created Excel report info and
 * can also responsible for the further processing of that file.
 */
public interface JxlsOutput {
	
    /**
     * @return new OutputStream where Jxls writes the created Excel report into
     * @throws IOException
     */
	OutputStream getOutputStream() throws IOException;
}
