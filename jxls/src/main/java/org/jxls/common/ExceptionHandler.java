package org.jxls.common;

public interface ExceptionHandler {

    void handleEvaluationException(Exception e, String cell, String expression);

    void handleFormulaException(Exception e, String cell, String formula);
    
    void handleCellException(Exception e, String cell, String contextKeys);
    
    void handleTransformException(Exception e, String sourceCell, String targetCell);
    
    void handleUpdateRowHeightsException(Exception e, int sourceRow, int targetRow);
}
