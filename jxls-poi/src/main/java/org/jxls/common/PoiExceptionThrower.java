package org.jxls.common;

import org.jxls.logging.JxlsLogger;

public class PoiExceptionThrower implements JxlsLogger {

    @Override
    public void handleCellException(Exception e, String cell, Context context) {
        throw new JxlsException("Failed to write a cell with " + cell, e);
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

    @Override
    public void handleGetObjectPropertyException(Exception e, Object obj, String propertyName) {
        throw new JxlsException("Failed to get property '" + propertyName + "' of object " + obj, e);
    }

    @Override
    public void handleSetObjectPropertyException(Exception e, Object obj, String propertyName, String propertyValue) {
        throw new JxlsException("Failed to set property '" + propertyName + "' to value '" + propertyValue + "' for object " + obj, e);
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
        throw new JxlsException(msg);
    }
    
    @Override
    public void error(Throwable e, String msg) {
        throw new JxlsException(msg, e);
    }
    
    protected void write(String level, String msg, Throwable e) {
        System.err.println("JXLS [" + level + "] " + msg);
        if (e != null) {
            e.printStackTrace();
        }
    }
}
