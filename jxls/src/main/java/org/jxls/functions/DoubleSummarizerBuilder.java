package org.jxls.functions;

public class DoubleSummarizerBuilder implements SummarizerBuilder<Double> {

    @Override
    public Summarizer<Double> build() {
        return new Summarizer<Double>() {
            private double sum = 0;

            @Override
            public void add(Object number) {
                if (number != null) {
                    sum += ((Double) number).doubleValue();
                }
            }

            @Override
            public Double getSum() {
                return Double.valueOf(sum);
            }
        };
    }
}
