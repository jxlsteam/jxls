package com.jxls.writer.transform.poi;

import com.jxls.writer.Cell;
import com.jxls.writer.CellData;
import com.jxls.writer.Pos;
import com.jxls.writer.command.Context;
import com.jxls.writer.transform.Transformer;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

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

    public List<Pos> getTargetCells(int sheet, int row, int col) {
        return new ArrayList<Pos>();
    }

    public void transform(Cell pos, Cell newCell, Context context) {
        if(cellData == null ||  cellData.length <= pos.getSheetIndex() || cellData[pos.getSheetIndex()] == null ||
                cellData[pos.getSheetIndex()].length <= pos.getRow() || cellData[pos.getSheetIndex()][pos.getRow()] == null ||
                cellData[pos.getSheetIndex()][pos.getRow()].length <= pos.getCol()) return;
        PoiCellData cellData = this.cellData[pos.getSheetIndex()][pos.getRow()][pos.getCol()];
        if(cellData != null){
            int numberOfSheets = workbook.getNumberOfSheets();
            while(numberOfSheets <= newCell.getSheetIndex() ){
                workbook.createSheet();
                numberOfSheets = workbook.getNumberOfSheets();
            }
            Sheet destSheet = workbook.getSheetAt(newCell.getSheetIndex());
            if(!ignoreColumnProps){
                destSheet.setColumnWidth(newCell.getCol(), sheetData[pos.getSheetIndex()].getColumnWidth(pos.getCol()));
            }
            Row destRow = destSheet.getRow(newCell.getRow());
            if (destRow == null) {
                destRow = destSheet.createRow(newCell.getRow());
            }
            if(!ignoreRowProps){
                destSheet.getRow(newCell.getRow()).setHeight( sheetData[pos.getSheetIndex()].getRowData(pos.getRow()).getHeight());
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

    public void updateFormulaCell(Cell cell, String formulaString) {
        Sheet sheet = workbook.getSheetAt(cell.getSheetIndex());
        if( sheet == null){
            sheet = workbook.createSheet();
        }
        Row row = sheet.getRow(cell.getRow());
        if( row == null ){
            row = sheet.createRow(cell.getRow());
        }
        org.apache.poi.ss.usermodel.Cell poiCell = row.getCell(cell.getCol());
        if( poiCell == null ){
            poiCell = row.createCell(cell.getCol());
        }
        poiCell.setCellFormula( formulaString );
    }

    public Set<CellData> getFormulaCells() {
        Set<CellData> formulaCells = new HashSet<CellData>();
        for(int sheetInd = 0; sheetInd < cellData.length; sheetInd++){
            for(int rowInd = 0; rowInd < cellData[sheetInd].length; rowInd++){
                for(int colInd = 0; colInd < cellData[sheetInd][rowInd].length; colInd++){
                    if(cellData[sheetInd][rowInd][colInd]!= null && cellData[sheetInd][rowInd][colInd].isFormulaCell() ){
                        formulaCells.add(cellData[sheetInd][rowInd][colInd]);
                    }
                }
            }
        }
        return formulaCells;
    }
}
