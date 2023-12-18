package org.jxls.logging;

public class NoOpLogger implements JxlsLogger {

    @Override
    public void handleEvaluationException(Exception e, String cell, String expression) {
    }

    @Override
    public void handleFormulaException(Exception e, String cell, String formula) {
    }

    @Override
    public void handleCellException(Exception e, String cell, String contextKeys) {
    }

    @Override
    public void handleTransformException(Exception e, String sourceCell, String targetCell) {
    }

    @Override
    public void handleUpdateRowHeightsException(Exception e, int sourceRow, int targetRow) {
    }

    @Override
    public void handleSetObjectPropertyException(Exception e, Object obj, String propertyName, String propertyValue) {
    }

    @Override
    public void handleGetObjectPropertyException(Exception e, Object obj, String propertyName) {
    }
}
