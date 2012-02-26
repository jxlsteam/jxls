package com.jxls.writer.transform.poi;

import com.jxls.writer.common.CellData;
import com.jxls.writer.common.CellRef;
import com.jxls.writer.common.Context;
import com.jxls.writer.util.Util;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn
 *         Date: 1/23/12 6:39 PM
 */
public class PoiCellData extends CellData {
    static Logger logger = LoggerFactory.getLogger(PoiCellData.class);

    RichTextString richTextString;
    private CellStyle cellStyle;
    private Hyperlink hyperlink;
    private Comment comment;

    public PoiCellData(CellRef cellRef) {
        super(cellRef);
    }

    public static PoiCellData createCellData(CellRef cellRef, Cell cell){
        PoiCellData cellData = new PoiCellData(cellRef);
        cellData.readCell(cell);
        cellData.updateFormulaValue();
        return cellData;
    }

    public void readCell(Cell cell){
        readCellGeneralInfo(cell);
        readCellContents(cell);
        readCellStyle(cell);
    }

    private void readCellGeneralInfo(Cell cell) {
        hyperlink = cell.getHyperlink();
        comment = cell.getCellComment();
        if( comment != null && comment.getString() != null ){
            setCellComment( comment.getString().getString() );
        }
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    private void readCellContents(Cell cell) {
        switch( cell.getCellType() ){
            case Cell.CELL_TYPE_STRING:
                richTextString = cell.getRichStringCellValue();
                evaluationResult = richTextString.getString();
                cellValue = richTextString.getString();
                cellType = CellType.STRING;
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                evaluationResult = cell.getBooleanCellValue();
                cellValue = evaluationResult;
                cellType = CellType.BOOLEAN;
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if(DateUtil.isCellDateFormatted(cell)) {
                    evaluationResult = cell.getDateCellValue();
                    cellValue = evaluationResult;
                    cellType = CellType.DATE;
                } else {
                    evaluationResult = cell.getNumericCellValue();
                    cellValue = evaluationResult;
                    cellType = CellType.NUMBER;
                }
                break;
            case Cell.CELL_TYPE_FORMULA:
                formula = cell.getCellFormula();
                evaluationResult = formula;
                cellValue = formula;
                cellType = CellType.FORMULA;
                break;
            case Cell.CELL_TYPE_ERROR:
                evaluationResult = cell.getErrorCellValue();
                cellValue = evaluationResult;
                cellType = CellType.ERROR;
                break;
            case Cell.CELL_TYPE_BLANK:
                evaluationResult = null;
                cellValue = null;
                cellType = CellType.BLANK;
                break;
        }
    }

    private void readCellStyle(Cell cell) {
        cellStyle = cell.getCellStyle();
    }

    public void writeToCell(Cell cell, Context context){
        evaluate(context);
        updateCellGeneralInfo(cell);
        updateCellContents( cell );
        updateCellStyle( cell );
    }

    private void updateCellGeneralInfo(Cell cell) {
        cell.setCellType( getPoiCellType(targetCellType) );
        if( hyperlink != null ){
            cell.setHyperlink( hyperlink );
        }
        if(comment != null ){
            cell.setCellComment(comment);
        }
    }
    
    static int getPoiCellType(CellType cellType){
        if( cellType == null ){
            return Cell.CELL_TYPE_BLANK;
        }
        switch (cellType){
            case STRING: 
                return Cell.CELL_TYPE_STRING;
            case BOOLEAN: 
                return Cell.CELL_TYPE_BOOLEAN;
            case NUMBER:
            case DATE:
                return Cell.CELL_TYPE_NUMERIC;
            case FORMULA:
                return Cell.CELL_TYPE_FORMULA;
            case ERROR:
                return Cell.CELL_TYPE_ERROR;
            case BLANK:
                return Cell.CELL_TYPE_BLANK;
            default:
                return Cell.CELL_TYPE_BLANK;
        }
    }

    private void updateCellContents(Cell cell) {
        switch( targetCellType ){
            case STRING:
                cell.setCellValue((String) evaluationResult);
                break;
            case BOOLEAN:
                cell.setCellValue( (Boolean)evaluationResult );
                break;
            case NUMBER:
                    if( evaluationResult instanceof Integer){
                        cell.setCellValue(((Integer)evaluationResult).doubleValue());
                    }else{
                        cell.setCellValue((Double) evaluationResult);
                    }
                break;
            case FORMULA:
                try{
                    if( Util.formulaContainsJointedCellRef((String) evaluationResult) ){
                        cell.setCellValue((String)evaluationResult);
                    }else{
                        cell.setCellFormula((String) evaluationResult);
                    }
                }catch(FormulaParseException e){
                    String formulaString = "";
                    try{
                        formulaString = evaluationResult.toString();
                        logger.error("Failed to set cell formula " + formulaString + " for cell " + this.toString(), e);
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                        cell.setCellValue(formulaString);
                    }catch(Exception ex){
                        logger.warn("Failed to convert formula to string for cell " + this.toString());
                    }
                }
                break;
            case ERROR:
                cell.setCellErrorValue((Byte) evaluationResult);
                break;
        }
    }

    private void updateCellStyle(Cell cell) {
        cell.setCellStyle(cellStyle);
    }

}
