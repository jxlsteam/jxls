package org.jxls.common;

import org.jxls.logging.JxlsLogger;

public class PoiExceptionLogger implements JxlsLogger {

    @Override
    public void handleCellException(Exception e, String cell, String contextKeys) {
        error(e, "Failed to write a cell with " + cell + " and context keys " + contextKeys);
    }

    @Override
    public void handleFormulaException(Exception e, String cell, String formula) {
        error(e, "Failed to set formula = " + formula + " into cell = " + cell);
    }

    @Override
    public void handleTransformException(Exception e, String sourceCell, String targetCell) {
        error(e, "Failed to transform " + sourceCell + " into " + targetCell);
    }

    @Override
    public void handleUpdateRowHeightsException(Exception e, int sourceRow, int targetRow) {
        error(e, "Failed to update row height for src row=" + sourceRow + " and target row=" + targetRow);
    }

    @Override
    public void handleEvaluationException(Exception e, String cell, String expression) {
        warn(e, "Failed to evaluate collection expression " + expression);
    }

    @Override
    public void handleGetObjectPropertyException(Exception e, Object obj, String propertyName) {
        warn(e, "Failed to get property '" + propertyName + "' of object " + obj);
    }

    @Override
    public void handleSetObjectPropertyException(Exception e, Object obj, String propertyName, String propertyValue) {
        warn(e, "Failed to set property '" + propertyName + "' to value '" + propertyValue + "' for object " + obj);
    }
}
