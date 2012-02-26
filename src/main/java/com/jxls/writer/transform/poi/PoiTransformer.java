package com.jxls.writer.transform.poi;

import com.jxls.writer.common.CellData;
import com.jxls.writer.common.CellRef;
import com.jxls.writer.common.Context;
import com.jxls.writer.transform.AbstractTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * @author Leonid Vysochyn
 *         Date: 1/23/12 2:36 PM
 */
public class PoiTransformer extends AbstractTransformer {
    static Logger logger = LoggerFactory.getLogger(PoiTransformer.class);

    Workbook workbook;

    private PoiTransformer(Workbook workbook) {
        this.workbook = workbook;
    }
    
    public static PoiTransformer createTransformer(InputStream is) throws IOException, InvalidFormatException {
        Workbook workbook = WorkbookFactory.create(is);
        return createTransformer(workbook);
    }

    public static PoiTransformer createTransformer(Workbook workbook) {
        PoiTransformer transformer = new PoiTransformer(workbook);
        transformer.readCellData();
        return transformer;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    private void readCellData(){
        int numberOfSheets = workbook.getNumberOfSheets();
        for(int i = 0; i < numberOfSheets; i++){
            Sheet sheet = workbook.getSheetAt(i);
            SheetData sheetData = SheetData.createSheetData(sheet);
            sheetMap.put(sheetData.getSheetName(), sheetData);
        }
    }

    public void transform(CellRef srcCellRef, CellRef targetCellRef, Context context) {
        CellData cellData = this.getCellData(srcCellRef);
        if(cellData != null){
            cellData.addTargetPos(targetCellRef);
            if(targetCellRef == null || targetCellRef.getSheetName() == null){
                logger.info("Target cellRef is null or has empty sheet name, cellRef=" + targetCellRef);
                return;
            }
            Sheet destSheet = workbook.getSheet(targetCellRef.getSheetName());
            if(destSheet == null){
                destSheet = workbook.createSheet(targetCellRef.getSheetName());
            }
            SheetData sheetData = sheetMap.get(srcCellRef.getSheetName());
            if(!isIgnoreColumnProps()){
                destSheet.setColumnWidth(targetCellRef.getCol(), sheetData.getColumnWidth(srcCellRef.getCol()));
            }
            Row destRow = destSheet.getRow(targetCellRef.getRow());
            if (destRow == null) {
                destRow = destSheet.createRow(targetCellRef.getRow());
            }
            if(!isIgnoreRowProps()){
                destSheet.getRow(targetCellRef.getRow()).setHeight( sheetData.getRowData(srcCellRef.getRow()).getHeight());
            }
            org.apache.poi.ss.usermodel.Cell destCell = destRow.getCell(targetCellRef.getCol());
            if (destCell == null) {
                destCell = destRow.createCell(targetCellRef.getCol());
            }
            try{
                destCell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK);
                ((PoiCellData)cellData).writeToCell(destCell, context);
                copyMergedRegions(cellData, targetCellRef);
            }catch(Exception e){
                logger.error("Failed to write a cell with " + cellData + " and " + context, e);
            }
        }
    }

    private void copyMergedRegions(CellData sourceCellData, CellRef destCell) {
        if(sourceCellData.getSheetName() == null ){ throw new IllegalArgumentException("Sheet name is null in copyMergedRegion");}
        SheetData sheetData = sheetMap.get( sourceCellData.getSheetName() );
        CellRangeAddress cellMergedRegion = null;
        for (CellRangeAddress mergedRegion : sheetData.getMergedRegions()) {
            if(mergedRegion.getFirstRow() == sourceCellData.getRow() && mergedRegion.getFirstColumn() == sourceCellData.getCol()){
                cellMergedRegion = mergedRegion;
                break;
            }
        }
        if( cellMergedRegion != null){
            findAndRemoveExistingCellRegion(destCell);
            Sheet destSheet = workbook.getSheet(destCell.getSheetName());
            destSheet.addMergedRegion(new CellRangeAddress(destCell.getRow(), destCell.getRow() + cellMergedRegion.getLastRow() - cellMergedRegion.getFirstRow(),
                    destCell.getCol(), destCell.getCol() + cellMergedRegion.getLastColumn() - cellMergedRegion.getFirstColumn()));
        }
    }

    private void findAndRemoveExistingCellRegion(CellRef cellRef) {
        Sheet destSheet = workbook.getSheet(cellRef.getSheetName());
        int numMergedRegions = destSheet.getNumMergedRegions();
        List<Integer> regionsToRemove = new ArrayList<Integer>();
        for(int i = 0; i < numMergedRegions; i++){
            CellRangeAddress mergedRegion = destSheet.getMergedRegion(i);
            if( mergedRegion.getFirstRow() <= cellRef.getRow() && mergedRegion.getLastRow() >= cellRef.getRow() &&
                    mergedRegion.getFirstColumn() <= cellRef.getCol() && mergedRegion.getLastColumn() >= cellRef.getCol() ){
                destSheet.removeMergedRegion(i);
                break;
            }
        }
    }

    public void setFormula(CellRef cellRef, String formulaString) {
        if(cellRef == null || cellRef.getSheetName() == null ) return;
        Sheet sheet = workbook.getSheet(cellRef.getSheetName());
        if( sheet == null){
            sheet = workbook.createSheet(cellRef.getSheetName());
        }
        Row row = sheet.getRow(cellRef.getRow());
        if( row == null ){
            row = sheet.createRow(cellRef.getRow());
        }
        org.apache.poi.ss.usermodel.Cell poiCell = row.getCell(cellRef.getCol());
        if( poiCell == null ){
            poiCell = row.createCell(cellRef.getCol());
        }
        try{
            poiCell.setCellFormula( formulaString );
        }catch (Exception e){
            logger.error("Failed to set formula = " + formulaString + " into cell = " + cellRef.getCellName(), e);
        }
    }

    public void clearCell(CellRef cellRef) {
        if(cellRef == null || cellRef.getSheetName() == null ) return;
        Sheet sheet = workbook.getSheet(cellRef.getSheetName());
        if( sheet == null ) return;
        Row row = sheet.getRow(cellRef.getRow());
        if( row == null ) return;
        Cell cell = row.getCell(cellRef.getCol());
        if( cell == null ) return;
        cell.setCellType(Cell.CELL_TYPE_BLANK);
        cell.setCellStyle(workbook.getCellStyleAt((short) 0));
    }

    public List<CellData> getCommentedCells() {
        List<CellData> commentedCells = new ArrayList<CellData>();
        for (SheetData sheetData : sheetMap.values()) {
            for (RowData rowData : sheetData) {
                for (CellData cellData : rowData) {
                    if(cellData != null && cellData.getCellComment() != null ){
                        commentedCells.add(cellData);
                    }
                }
            }
        }
        return commentedCells;
    }

}
