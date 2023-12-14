package org.jxls.builder;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class JxlsOutputFile implements JxlsOutput {
	private final File file;
	private FileOutputStream fos;
	
	public JxlsOutputFile(File output) {
		file = output;
	}
	
	@Override
	public OutputStream getOutputStream() throws IOException {
		if (fos == null) {
			fos = new FileOutputStream(file);
		}
		return fos;
	}
}
