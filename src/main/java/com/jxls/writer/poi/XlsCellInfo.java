package com.jxls.writer.poi;

import org.apache.poi.ss.usermodel.*;

import java.util.Date;

/**
 * @author Leonid Vysochyn
 *         Date: 21.04.2009
 */
public class XlsCellInfo {
    RichTextString richTextString;
    boolean booleanValue;

    int cellType;
    private Date dateValue;
    private double doubleValue;
    private String formula;
    private CellStyle style;
    private Hyperlink hyperlink;
    private byte errorValue;
    private int rowIndex;
    private int colIndex;

    public int getColIndex() {
        return colIndex;
    }

    public void readCell(Cell cell){
        readCellGeneralInfo(cell);
        readCellContents(cell);
        readCellStyle(cell);
    }

    private void readCellGeneralInfo(Cell cell) {
        cellType = cell.getCellType();
        hyperlink = cell.getHyperlink();
        colIndex = cell.getColumnIndex();
        rowIndex = cell.getRowIndex();
    }

    private void readCellContents(Cell cell) {
        switch( cell.getCellType() ){
            case Cell.CELL_TYPE_STRING:
                richTextString = cell.getRichStringCellValue();
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                booleanValue = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if(DateUtil.isCellDateFormatted(cell)) {
                  dateValue = cell.getDateCellValue();
                } else {
                  doubleValue = cell.getNumericCellValue();
                }
                break;
            case Cell.CELL_TYPE_FORMULA:
                formula = cell.getCellFormula();
                break;
            case Cell.CELL_TYPE_ERROR:
                errorValue = cell.getErrorCellValue();
                break;
        }
    }

    private void readCellStyle(Cell cell) {
        style = cell.getCellStyle();
    }

    public void writeToCell(Cell cell){
        updateCellGeneralInfo(cell);
        updateCellContents( cell );
        updateCellStyle( cell );
    }

    private void updateCellGeneralInfo(Cell cell) {
        cell.setCellType( cellType );
        if( hyperlink != null ){
            cell.setHyperlink( hyperlink );
        }
    }

    private void updateCellContents(Cell cell) {
        switch( cellType ){
            case Cell.CELL_TYPE_STRING:
                cell.setCellValue( richTextString );
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                cell.setCellValue( booleanValue );
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if( dateValue != null ){
                    cell.setCellValue( dateValue );
                }else{
                    cell.setCellValue( doubleValue );
                }
                break;
            case Cell.CELL_TYPE_FORMULA:
                cell.setCellFormula(formula);
                break;
            case Cell.CELL_TYPE_ERROR:
                cell.setCellErrorValue( errorValue );
                break;
        }
    }

    private void updateCellStyle(Cell cell) {
        cell.setCellStyle( style );
    }

}
