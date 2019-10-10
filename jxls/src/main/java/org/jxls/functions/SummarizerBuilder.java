package org.jxls.functions;

public interface SummarizerBuilder<T> {

    /**
     * @return new Summarizer instance with initial value (usually zero)
     */
    Summarizer<T> build();
}
