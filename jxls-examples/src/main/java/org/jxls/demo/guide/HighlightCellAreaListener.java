package org.jxls.demo.guide;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jxls.common.AreaListener;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.poi.PoiTransformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn
 *         Date: 2/16/12 6:07 PM
 */
public class HighlightCellAreaListener implements AreaListener {
    static Logger logger = LoggerFactory.getLogger(HighlightCellAreaListener.class);

    PoiTransformer transformer;
    private final CellRef paymentCell = new CellRef("Template!C4");

    public HighlightCellAreaListener(PoiTransformer transformer) {
        this.transformer = transformer;
    }

    public void beforeApplyAtCell(CellRef cellRef, Context context) {

    }

    public void afterApplyAtCell(CellRef cellRef, Context context) {

    }

    public void beforeTransformCell(CellRef srcCell, CellRef targetCell, Context context) {

    }

    public void afterTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
        System.out.println("Source: " + srcCell.getCellName() + ", Target: " + targetCell.getCellName());
        if(paymentCell.equals(srcCell)){ // we are at employee payment cell
            Employee employee = (Employee) context.getVar("employee");
            if( employee.getPayment().doubleValue() > 2000 ){ // highlight payment when >= $2000
                logger.info("highlighting payment for employee " + employee.getName());
                highlightCell(targetCell);
            }
        }
    }

    private void highlightCell(CellRef cellRef) {
        Workbook workbook = transformer.getWorkbook();
        Sheet sheet = workbook.getSheet(cellRef.getSheetName());
        Cell cell = sheet.getRow(cellRef.getRow()).getCell(cellRef.getCol());
        CellStyle cellStyle = cell.getCellStyle();
        CellStyle newCellStyle = workbook.createCellStyle();
        newCellStyle.setDataFormat( cellStyle.getDataFormat() );
        newCellStyle.setFont( workbook.getFontAt( cellStyle.getFontIndex() ));
        newCellStyle.setFillBackgroundColor( cellStyle.getFillBackgroundColor());
        newCellStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(newCellStyle);
    }
}
