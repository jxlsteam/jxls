package org.jxls.demo;

import org.jxls.area.XlsArea;
import org.jxls.common.AreaListener;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.demo.model.Employee;
import org.jxls.transform.poi.PoiTransformer;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn
 *         Date: 2/16/12 6:07 PM
 */
public class SimpleAreaListener implements AreaListener {
    static Logger logger = LoggerFactory.getLogger(SimpleAreaListener.class);
    
    XlsArea area;
    PoiTransformer transformer;
    private final CellRef bonusCell1 = new CellRef("Template!E9");
    private final CellRef bonusCell2 =new CellRef("Template!E18");

    public SimpleAreaListener(XlsArea area) {
        this.area = area;
        transformer = (PoiTransformer) area.getTransformer();
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
        Workbook workbook = transformer.getWorkbook();
        Sheet sheet = workbook.getSheet(cellRef.getSheetName());
        Cell cell = sheet.getRow(cellRef.getRow()).getCell(cellRef.getCol());
        CellStyle cellStyle = cell.getCellStyle();
        CellStyle newCellStyle = workbook.createCellStyle();
        newCellStyle.setDataFormat( cellStyle.getDataFormat() );
        newCellStyle.setFont( workbook.getFontAt( cellStyle.getFontIndex() ));
        newCellStyle.setFillBackgroundColor( cellStyle.getFillBackgroundColor());
        newCellStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        //newCellStyle.setFillForegroundColor( cellStyle.getFillForegroundColor());
        newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(newCellStyle);
    }
}
