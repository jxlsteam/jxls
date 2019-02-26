package org.jxls.demo;

import org.jxls.area.XlsArea;
import org.jxls.common.AreaListener;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.demo.model.Employee;
import org.jxls.transform.jexcel.JexcelTransformer;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.write.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn
 *         Date: 11/15/13
 */
public class JexcelSimpleAreaListener implements AreaListener {
    static Logger logger = LoggerFactory.getLogger(JexcelSimpleAreaListener.class);

    XlsArea area;
    JexcelTransformer transformer;
    private final CellRef bonusCell1 = new CellRef("Template!E9");
    private final CellRef bonusCell2 =new CellRef("Template!E18");

    public JexcelSimpleAreaListener(XlsArea area) {
        this.area = area;
        transformer = (JexcelTransformer) area.getTransformer();
    }

    public void beforeApplyAtCell(CellRef cellRef, Context context) {

    }

    public void afterApplyAtCell(CellRef cellRef, Context context) {

    }

    public void beforeTransformCell(CellRef srcCell, CellRef targetCell, Context context) {

    }

    public void afterTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
        if(bonusCell1.equals(srcCell) || bonusCell2.equals(srcCell)){ // we are at employee bonus cell
            Employee employee = (Employee) context.getVar("employee");
            if( employee.getBonus() >= 0.2 ){ // highlight bonus when >= 20%
                logger.info("highlighting bonus for employee " + employee.getName());
                highlightBonus(targetCell);
            }
        }
    }

    private void highlightBonus(CellRef cellRef) {
        WritableWorkbook workbook = transformer.getWritableWorkbook();
        WritableSheet sheet = workbook.getSheet(cellRef.getSheetName());
        WritableCell cell = sheet.getWritableCell(cellRef.getCol(), cellRef.getRow());
        CellFormat currentFormat = cell.getCellFormat();
        WritableCellFormat newFormat = new WritableCellFormat(currentFormat);
        Colour highlightColor = Colour.ORANGE;
        try {
            newFormat.setBackground(highlightColor);
        } catch (WriteException e) {
            logger.error("Failed to set background color to cell " + cellRef, e);
        }
        cell.setCellFormat( newFormat );
    }
}
