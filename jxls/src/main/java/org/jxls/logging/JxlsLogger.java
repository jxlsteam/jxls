package org.jxls.logging;

public interface JxlsLogger {

    // special logging
    
    void handleEvaluationException(Exception e, String cell, String expression);
    void handleFormulaException(Exception e, String cell, String formula);
    void handleCellException(Exception e, String cell, String contextKeys);
    void handleTransformException(Exception e, String sourceCell, String targetCell);
    void handleUpdateRowHeightsException(Exception e, int sourceRow, int targetRow);
    void handleSetObjectPropertyException(Exception e, Object obj, String propertyName, String propertyValue);
    void handleGetObjectPropertyException(Exception e, Object obj, String propertyName);

    // generic logging
    
    default void debug(String msg) {}
    
    default void info(String msg) {}

    default void warn(String msg) {}
    default void warn(Throwable e, String msg) {}
    
    default void error(String msg) {}
    default void error(Throwable e, String msg) {}
}
