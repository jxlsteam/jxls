package org.jxls.transform.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jxls.builder.JxlsStreaming;
import org.jxls.logging.JxlsLogger;
import org.jxls.transform.TemplateProcessor;

/**
 * POI bug 66679 compensation
 */
public class CellStyleGeneralEnsurer implements TemplateProcessor {

    @Override
    public void process(Object template, JxlsStreaming streaming, JxlsLogger logger) {
        if (template instanceof XSSFWorkbook workbook) {
            process(workbook);
        }
    }
    
    public void process(XSSFWorkbook workbook) {
        // iterate over all cells
        for (Sheet sheet : workbook) {
            for (Row row : sheet) {
                for (Cell cell : row) {
                    processCell((XSSFCell) cell);
                }
            }
        }
    }

    protected void processCell(XSSFCell cell) {
        if (cell.getCTCell().isSetS()) {
            return; // cell is ok
        }
        String style = cell.getSheet().getColumnStyle(cell.getColumnIndex()).getDataFormatString();
        if ("General".equals(style)) {
            return; // Default column style is the same, so a fix is not necessary.
        }
        String content = cell.getStringCellValue();
        if (content != null && content.contains("${")) { // process only Jxls relevant cells
            fix(cell);
        }
    }

    /**
     * Overwrite method if you just want to collect the errors for an unit test.
     * @param cell has missing cell style 'General'
     */
    protected void fix(XSSFCell cell) {
        cell.setCellStyle(cell.getRow().getSheet().getWorkbook().getCellStyleAt(0));
    }
}
