package org.jxls.common;

public class EvaluationResult {
    private final Object result;
    private final boolean setTargetCellType;

    public EvaluationResult(Object result) {
        this(result, false);
    }

    public EvaluationResult(Object result, boolean setTargetCellType) {
        this.result = result;
        this.setTargetCellType = setTargetCellType;
    }

    public Object getResult() {
        return result;
    }

    public boolean isSetTargetCellType() {
        return setTargetCellType;
    }
}
