package com.jxls.writer.transform.poi;

import com.jxls.writer.Cell;
import com.jxls.writer.command.Context;
import com.jxls.writer.transform.Transformer;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn
 *         Date: 1/23/12 2:36 PM
 */
public class PoiTransformer implements Transformer {
    static Logger logger = LoggerFactory.getLogger(PoiTransformer.class);

    Workbook workbook;
    CellData[][][] cellDatas;

    public PoiTransformer(Workbook workbook) {
        this.workbook = workbook;
        readCellData();
    }
    
    private void readCellData(){
        int numberOfSheets = workbook.getNumberOfSheets();
        cellDatas = new CellData[numberOfSheets][][];
        for(int sheetIndex = 0; sheetIndex < numberOfSheets; sheetIndex++){
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            int numberOfRows = sheet.getLastRowNum() + 1;
            cellDatas[sheetIndex] = new CellData[numberOfRows][];
            for(int rowIndex = 0; rowIndex < numberOfRows; rowIndex++){
                Row row = sheet.getRow(rowIndex);
                if( row != null ){
                    int numberOfCells = row.getLastCellNum() + 1;
                    cellDatas[sheetIndex][rowIndex] = new CellData[numberOfCells];
                    for(int cellIndex = 0; cellIndex < numberOfCells; cellIndex++){
                        org.apache.poi.ss.usermodel.Cell cell = row.getCell(cellIndex);
                        if(cell != null ){
                            cellDatas[sheetIndex][rowIndex][cellIndex] = CellData.createCellData(cell);
                        }
                    }
                }
            }
        }
    }
    
    public CellData getCellData(int col, int row, int sheet){
        return cellDatas[sheet][row][col];
    }

    public void transform(Cell pos, Cell newCell, Context context) {
        if(cellDatas == null ||  cellDatas.length <= pos.getSheetIndex() || cellDatas[pos.getSheetIndex()] == null ||
                cellDatas[pos.getSheetIndex()].length <= pos.getRow() || cellDatas[pos.getSheetIndex()][pos.getRow()] == null ||
                cellDatas[pos.getSheetIndex()][pos.getRow()].length <= pos.getCol()) return;
        CellData cellData = cellDatas[pos.getSheetIndex()][pos.getRow()][pos.getCol()];
        if(cellData != null){
            int numberOfSheets = workbook.getNumberOfSheets();
            while(numberOfSheets <= newCell.getSheetIndex() ){
                workbook.createSheet();
                numberOfSheets = workbook.getNumberOfSheets();
            }
            Sheet destSheet = workbook.getSheetAt(newCell.getSheetIndex());
            Row destRow = destSheet.getRow(newCell.getRow());
            if (destRow == null) {
                destRow = destSheet.createRow(newCell.getRow());
            }
            org.apache.poi.ss.usermodel.Cell destCell = destRow.getCell(newCell.getCol());
            if (destCell == null) {
                destCell = destRow.createCell(newCell.getCol());
            }
            try{
                destCell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK);
                cellData.writeToCell(destCell, context);
            }catch(Exception e){
                logger.error("Failed to write a cell with " + cellData + " and " + context, e);
            }
        }
    }
}
