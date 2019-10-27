package org.jxls.functions;

public interface Summarizer<T> {
    
    /**
     * Casts number to type T and adds value to sum if number not null.
     * @param number a number to add
     */
    void add(Object number);
    
    /**
     * @return sum of type T
     */
    T getSum();
}
