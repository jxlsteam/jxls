package com.jxls.writer.transform.poi;

import com.jxls.writer.CellData;
import com.jxls.writer.Pos;
import com.jxls.writer.command.Context;
import com.jxls.writer.transform.AbstractTransformer;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

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

    public static PoiTransformer createTransformer(Workbook workbook) {
        PoiTransformer transformer = new PoiTransformer(workbook);
        transformer.readCellData();
        return transformer;
    }

    private void readCellData(){
        int numberOfSheets = workbook.getNumberOfSheets();
        for(int i = 0; i < numberOfSheets; i++){
            Sheet sheet = workbook.getSheetAt(i);
            SheetData sheetData = SheetData.createSheetData(sheet);
            sheetMap.put(sheetData.getSheetName(), sheetData);
        }
    }

    public void transform(Pos pos, Pos newPos, Context context) {
//        if(cellData == null ||  cellData.length <= pos.getSheet() || cellData[pos.getSheet()] == null ||
//                cellData[pos.getSheet()].length <= pos.getRow() || cellData[pos.getSheet()][pos.getRow()] == null ||
//                cellData[pos.getSheet()][pos.getRow()].length <= pos.getCol()) return;
        CellData cellData = this.getCellData(pos);
        if(cellData != null){
            cellData.addTargetPos(newPos);
            if(newPos == null || newPos.getSheetName() == null){
                logger.info("Target pos is null or has empty sheet name, pos=" + newPos);
                return;
            }
            Sheet destSheet = workbook.getSheet(newPos.getSheetName());
            if(destSheet == null){
                destSheet = workbook.createSheet(newPos.getSheetName());
            }
            SheetData sheetData = sheetMap.get(pos.getSheetName());
            if(!isIgnoreColumnProps()){
                destSheet.setColumnWidth(newPos.getCol(), sheetData.getColumnWidth(pos.getCol()));
            }
            Row destRow = destSheet.getRow(newPos.getRow());
            if (destRow == null) {
                destRow = destSheet.createRow(newPos.getRow());
            }
            if(!isIgnoreRowProps()){
                destSheet.getRow(newPos.getRow()).setHeight( sheetData.getRowData(pos.getRow()).getHeight());
            }
            org.apache.poi.ss.usermodel.Cell destCell = destRow.getCell(newPos.getCol());
            if (destCell == null) {
                destCell = destRow.createCell(newPos.getCol());
            }
            try{
                destCell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK);
                ((PoiCellData)cellData).writeToCell(destCell, context);
                copyMergedRegions(cellData, newPos);
            }catch(Exception e){
                logger.error("Failed to write a cell with " + cellData + " and " + context, e);
            }
        }
    }

    private void copyMergedRegions(CellData sourceCellData, Pos destCell) {
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
            Sheet destSheet = workbook.getSheetAt(destCell.getSheet());
            destSheet.addMergedRegion(new CellRangeAddress(destCell.getRow(), destCell.getRow() + cellMergedRegion.getLastRow() - cellMergedRegion.getFirstRow(),
                    destCell.getCol(), destCell.getCol() + cellMergedRegion.getLastColumn() - cellMergedRegion.getFirstColumn()));
        }
    }

    private void findAndRemoveExistingCellRegion(Pos pos) {
        Sheet destSheet = workbook.getSheetAt(pos.getSheet());
        int numMergedRegions = destSheet.getNumMergedRegions();
        List<Integer> regionsToRemove = new ArrayList<Integer>();
        for(int i = 0; i < numMergedRegions; i++){
            CellRangeAddress mergedRegion = destSheet.getMergedRegion(i);
            if( mergedRegion.getFirstRow() <= pos.getRow() && mergedRegion.getLastRow() >= pos.getRow() &&
                    mergedRegion.getFirstColumn() <= pos.getCol() && mergedRegion.getLastColumn() >= pos.getCol() ){
                destSheet.removeMergedRegion(i);
                break;
            }
        }
    }

    public void setFormula(Pos pos, String formulaString) {
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

    public int getSheetIndex(String sheetName) {
        return workbook.getSheetIndex(sheetName);
    }

}
