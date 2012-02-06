package com.jxls.writer.transform.poi;

import com.jxls.writer.CellData;
import com.jxls.writer.Pos;
import com.jxls.writer.command.Context;
import com.jxls.writer.transform.Transformer;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.*;

/**
 * @author Leonid Vysochyn
 *         Date: 1/23/12 2:36 PM
 */
public class PoiTransformer implements Transformer {
    static Logger logger = LoggerFactory.getLogger(PoiTransformer.class);

    Workbook workbook;
    PoiCellData[][][] cellData;
    SheetData[] sheetData;
    boolean ignoreColumnProps = false;
    boolean ignoreRowProps = false;

    public PoiTransformer(Workbook workbook) {
        this.workbook = workbook;
        readCellData();
    }

    public boolean isIgnoreColumnProps() {
        return ignoreColumnProps;
    }

    public void setIgnoreColumnProps(boolean ignoreColumnProps) {
        this.ignoreColumnProps = ignoreColumnProps;
    }

    public boolean isIgnoreRowProps() {
        return ignoreRowProps;
    }

    public void setIgnoreRowProps(boolean ignoreRowProps) {
        this.ignoreRowProps = ignoreRowProps;
    }

    private void readCellData(){
        int numberOfSheets = workbook.getNumberOfSheets();
        cellData = new PoiCellData[numberOfSheets][][];
        sheetData = new SheetData[numberOfSheets];
        for(int sheetIndex = 0; sheetIndex < numberOfSheets; sheetIndex++){
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            sheetData[sheetIndex] = SheetData.createSheetData(sheet);
            int numberOfRows = sheet.getLastRowNum() + 1;
            cellData[sheetIndex] = new PoiCellData[numberOfRows][];
            for(int rowIndex = 0; rowIndex < numberOfRows; rowIndex++){
                Row row = sheet.getRow(rowIndex);
                if( row != null ){
                    int numberOfCells = row.getLastCellNum() + 1;
                    cellData[sheetIndex][rowIndex] = new PoiCellData[numberOfCells];
                    for(int cellIndex = 0; cellIndex < numberOfCells; cellIndex++){
                        org.apache.poi.ss.usermodel.Cell cell = row.getCell(cellIndex);
                        if(cell != null ){
                            cellData[sheetIndex][rowIndex][cellIndex] = PoiCellData.createCellData(cell);
                        }
                    }
                }
            }
        }
    }
    
    public CellData getCellData(int sheet, int row, int col){
        return cellData[sheet][row][col];
    }

    public List<Pos> getTargetPos(int sheet, int row, int col) {
        CellData cellData = getCellData(sheet, row, col);
        if( cellData != null ){
            return cellData.getTargetPos();
        }else{
            return new ArrayList<Pos>();
        }
    }

    public void transform(Pos pos, Pos newPos, Context context) {
        if(cellData == null ||  cellData.length <= pos.getSheet() || cellData[pos.getSheet()] == null ||
                cellData[pos.getSheet()].length <= pos.getRow() || cellData[pos.getSheet()][pos.getRow()] == null ||
                cellData[pos.getSheet()][pos.getRow()].length <= pos.getCol()) return;
        PoiCellData cellData = this.cellData[pos.getSheet()][pos.getRow()][pos.getCol()];
        if(cellData != null){
            cellData.addTargetPos(newPos);
            int numberOfSheets = workbook.getNumberOfSheets();
            while(numberOfSheets <= newPos.getSheet() ){
                workbook.createSheet();
                numberOfSheets = workbook.getNumberOfSheets();
            }
            Sheet destSheet = workbook.getSheetAt(newPos.getSheet());
            if(!ignoreColumnProps){
                destSheet.setColumnWidth(newPos.getCol(), sheetData[pos.getSheet()].getColumnWidth(pos.getCol()));
            }
            Row destRow = destSheet.getRow(newPos.getRow());
            if (destRow == null) {
                destRow = destSheet.createRow(newPos.getRow());
            }
            if(!ignoreRowProps){
                destSheet.getRow(newPos.getRow()).setHeight( sheetData[pos.getSheet()].getRowData(pos.getRow()).getHeight());
            }
            org.apache.poi.ss.usermodel.Cell destCell = destRow.getCell(newPos.getCol());
            if (destCell == null) {
                destCell = destRow.createCell(newPos.getCol());
            }
            try{
                destCell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK);
                cellData.writeToCell(destCell, context);
            }catch(Exception e){
                logger.error("Failed to write a cell with " + cellData + " and " + context, e);
            }
        }
    }

    public void updateFormulaCell(Pos pos, String formulaString) {
        Sheet sheet = workbook.getSheetAt(pos.getSheet());
        if( sheet == null){
            sheet = workbook.createSheet();
        }
        Row row = sheet.getRow(pos.getRow());
        if( row == null ){
            row = sheet.createRow(pos.getRow());
        }
        org.apache.poi.ss.usermodel.Cell poiCell = row.getCell(pos.getCol());
        if( poiCell == null ){
            poiCell = row.createCell(pos.getCol());
        }
        try{
            poiCell.setCellFormula( formulaString );
        }catch (Exception e){
            logger.error("Failed to set formula = " + formulaString + " into cell = " + pos.getCellName(), e);
        }
    }

    public Set<CellData> getFormulaCells() {
        Set<CellData> formulaCells = new HashSet<CellData>();
        for(int sheetInd = 0; sheetInd < cellData.length; sheetInd++){
            for(int rowInd = 0; rowInd < cellData[sheetInd].length; rowInd++){
                if(cellData[sheetInd][rowInd] != null ){
                    for(int colInd = 0; colInd < cellData[sheetInd][rowInd].length; colInd++){
                        if(cellData[sheetInd][rowInd][colInd]!= null && cellData[sheetInd][rowInd][colInd].isFormulaCell() ){
                            formulaCells.add(cellData[sheetInd][rowInd][colInd]);
                        }
                    }
                }
            }
        }
        return formulaCells;
    }
}
