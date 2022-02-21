package org.jxls.common;

import org.jxls.area.XlsArea;
import org.jxls.command.EachCommand;
import org.jxls.transform.poi.PoiTransformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class PoiExceptionLogger implements ExceptionHandler {
    private static final Logger logger1 = LoggerFactory.getLogger(PoiTransformer.class);
    private static final Logger logger2 = LoggerFactory.getLogger(XlsArea.class);
    private static final Logger logger3 = LoggerFactory.getLogger(EachCommand.class);

    @Override
    public void handleCellException(Exception e, String cell, String contextKeys) {
        logger1.error("Failed to write a cell with {} and context keys {}", cell, contextKeys, e);
    }

    @Override
    public void handleFormulaException(Exception e, String cell, String formula) {
        logger1.error("Failed to set formula = " + formula + " into cell = " + cell, e);
    }

    @Override
    public void handleTransformException(Exception e, String sourceCell, String targetCell) {
        logger2.error("Failed to transform " + sourceCell + " into " + targetCell, e);
    }

    @Override
    public void handleUpdateRowHeightsException(Exception e, int sourceRow, int targetRow) {
        logger2.error("Failed to update row height for src row={} and target row={}", sourceRow, targetRow, e);
    }

    @Override
    public void handleEvaluationException(Exception e, String cell, String expression) {
        logger3.warn("Failed to evaluate collection expression {}", expression, e);
    }
}
