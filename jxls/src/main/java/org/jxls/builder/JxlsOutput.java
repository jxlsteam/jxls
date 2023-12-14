package org.jxls.builder;

import java.io.IOException;
import java.io.OutputStream;

public interface JxlsOutput {
	
	OutputStream getOutputStream() throws IOException;
}
