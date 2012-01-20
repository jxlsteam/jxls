package com.jxls.writer.expression;

/**
 * Date: Nov 2, 2009
 *
 * @author Leonid Vysochyn
 */
public class EvaluationException extends RuntimeException{

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
