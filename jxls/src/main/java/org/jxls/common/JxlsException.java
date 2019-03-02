package org.jxls.common;

/**
 * This is just a wrapper for non-runtime exceptions
 */
public class JxlsException extends RuntimeException {
    private static final long serialVersionUID = -3618771481378341613L;

    public JxlsException() {
    }

    public JxlsException(String message) {
        super(message);
    }

    public JxlsException(String message, Throwable cause) {
        super(message, cause);
    }

    public JxlsException(Throwable cause) {
        super(cause);
    }

    public JxlsException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}

