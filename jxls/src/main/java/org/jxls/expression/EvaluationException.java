package org.jxls.expression;

/**
 * Custom exception class for exceptions thrown during evaluation of expressions
 * 
 * @author Leonid Vysochyn
 */
public class EvaluationException extends RuntimeException {
    private static final long serialVersionUID = -611735067629358064L;

    public EvaluationException() {
    }

    public EvaluationException(String message) {
        super(message);
    }

    public EvaluationException(String message, Throwable cause) {
        super(message, cause);
    }

    public EvaluationException(Throwable cause) {
        super(cause);
    }
}
