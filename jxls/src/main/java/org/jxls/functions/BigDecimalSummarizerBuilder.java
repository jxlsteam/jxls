package org.jxls.functions;

import java.math.BigDecimal;

public class BigDecimalSummarizerBuilder implements SummarizerBuilder<BigDecimal> {

    @Override
    public Summarizer<BigDecimal> build() {
        return new Summarizer<BigDecimal>() {
            private BigDecimal sum = BigDecimal.ZERO;

            @Override
            public void add(Object number) {
                if (number != null) {
                    sum = sum.add((BigDecimal) number);
                }

            }

            @Override
            public BigDecimal getSum() {
                return sum;
            }
        };
    }
}
