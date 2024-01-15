package org.jxls.util;

import org.jxls.common.JxlsException;

public class JxlsPropertyException extends JxlsException {
	private final boolean write;
	
	public JxlsPropertyException(String msg, boolean write, Throwable e) {
		super(msg, e);
		this.write = write;
	}

	public boolean isWrite() {
		return write;
	}
	
	public boolean isRead() {
		return !write;
	}
}
