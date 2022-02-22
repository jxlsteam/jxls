package org.jxls.common;

import org.jxls.common.JxlsException;
import org.jxls.common.ExceptionHandler;

public class PoiExceptionThrower implements ExceptionHandler {

    @Override
    public void handleCellException(Exception e, String cell, String contextKeys) {
        throw new JxlsException("Failed to write a cell with " + cell + " and context keys " + contextKeys, e);
    }

    @Override
    public void handleFormulaException(Exception e, String cell, String formula) {
        throw new JxlsException("Failed to set formula \"" + formula + "\" into cell " + cell, e);
    }

    @Override
    public void handleTransformException(Exception e, String sourceCell, String targetCell) {
        throw new JxlsException("Failed to transform " + sourceCell + " into " + targetCell + "\n" + e.getMessage(), e);
    }

    @Override
    public void handleUpdateRowHeightsException(Exception e, int sourceRow, int targetRow) {
        throw new JxlsException("Failed to update row height for source row " + sourceRow + " and target row " + targetRow, e);
    }

    @Override
    public void handleEvaluationException(Exception e, String cell, String expression) {
        throw new JxlsException("Failed to evaluate collection expression \"" + expression + "\" in each command at " + cell, e);
    }
}
