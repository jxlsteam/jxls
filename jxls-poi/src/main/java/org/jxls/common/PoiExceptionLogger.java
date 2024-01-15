package org.jxls.common;

import org.jxls.logging.JxlsLogger;

public class PoiExceptionLogger implements JxlsLogger {

    @Override
    public void handleCellException(Exception e, String cell, Context context) {
        error(e, "Failed to write a cell with " + cell);
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
    
    @Override
    public void handleSheetNameChange(String invalidSheetName, String newSheetName) {
        info("Change invalid sheet name " + invalidSheetName + " to " + newSheetName);
    }

    @Override
    public void debug(String msg) {
    }

    @Override
    public void info(String msg) {
        System.out.println("JXLS [INFO] " + msg);
    }

    @Override
    public void warn(String msg) {
        write("WARN", msg, null);
    }

    @Override
    public void warn(Throwable e, String msg) {
        write("WARN", msg, e);
    }

    @Override
    public void error(String msg) {
        write("ERROR", msg, null);
    }

    @Override
    public void error(Throwable e, String msg) {
        write("ERROR", msg, e);
    }
    
    protected void write(String level, String msg, Throwable e) {
        System.err.println("JXLS [" + level + "] " + msg);
        if (e != null) {
            e.printStackTrace();
        }
    }
}
