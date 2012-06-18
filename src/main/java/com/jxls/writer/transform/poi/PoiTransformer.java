package com.jxls.writer.transform.poi;

import com.jxls.writer.common.*;
import com.jxls.writer.transform.AbstractTransformer;
import com.jxls.writer.common.RowData;
import com.jxls.writer.common.SheetData;
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
    public static final int MAX_COLUMN_TO_READ_COMMENT = 10; 
    
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
            SheetData sheetData = PoiSheetData.createSheetData(sheet);
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
        PoiSheetData sheetData = (PoiSheetData)sheetMap.get( sourceCellData.getSheetName() );
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
//  todo: after removing cell comment MS Excel displays warning 'File Error: data may have been lost' when opening the file  (seems like POI issue)
//        if( cell.getCellComment() != null ){
//            cell.removeCellComment();
//        }
    }

    public List<CellData> getCommentedCells() {
        List<CellData> commentedCells = new ArrayList<CellData>();
        for (SheetData sheetData : sheetMap.values()) {
            for (RowData rowData : sheetData) {
                if( rowData == null ) continue;
                for (CellData cellData : rowData) {
                    if(cellData != null && cellData.getCellComment() != null ){
                        commentedCells.add(cellData);
                    }
                }
                if( rowData.getNumberOfCells() == 0 ){
                    List<CellData> commentedCellData = readCommentsFromSheet(((PoiSheetData)sheetData).getSheet(), ((PoiRowData)rowData).getRow());
                    commentedCells.addAll( commentedCellData );
                }
            }
        }
        return commentedCells;
    }

    public void addImage(AreaRef areaRef, int imageIdx) {
        CreationHelper helper = workbook.getCreationHelper();
        Sheet sheet = workbook.getSheet(areaRef.getSheetName());
        Drawing drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(areaRef.getFirstCellRef().getCol());
        anchor.setRow1(areaRef.getFirstCellRef().getRow());
        anchor.setCol2(areaRef.getLastCellRef().getCol());
        anchor.setRow2(areaRef.getLastCellRef().getRow());
        drawing.createPicture(anchor, imageIdx);
    }

    public void addImage(AreaRef areaRef, byte[] imageBytes, ImageType imageType) {
        int poiPictureType = findPoiPictureTypeByImageType(imageType);
        int pictureIdx = workbook.addPicture(imageBytes, poiPictureType);
        addImage(areaRef, pictureIdx);
    }


    private int findPoiPictureTypeByImageType(ImageType imageType){
        int poiType = -1;
        if( imageType == null ){
            throw new IllegalArgumentException("Image type is undefined");
        }
        switch (imageType){
            case PNG:
                poiType = Workbook.PICTURE_TYPE_PNG;
                break;
            case JPEG:
                poiType = Workbook.PICTURE_TYPE_JPEG;
                break;
            case EMF:
                poiType = Workbook.PICTURE_TYPE_EMF;
                break;
            case WMF:
                poiType = Workbook.PICTURE_TYPE_WMF;
                break;
            case DIB:
                poiType = Workbook.PICTURE_TYPE_DIB;
                break;
            case PICT:
                poiType = Workbook.PICTURE_TYPE_PICT;
                break;
        }
        return poiType;
    }

    private List<CellData> readCommentsFromSheet(Sheet sheet, Row row) {
        List<CellData> commentDataCells = new ArrayList<CellData>();
        int rowNum = row.getRowNum();
        for(int i = 0; i <= MAX_COLUMN_TO_READ_COMMENT; i++){
            Comment comment = sheet.getCellComment(rowNum, i);
            if( comment != null && comment.getString() != null ){
                CellData cellData = new CellData( new CellRef(sheet.getSheetName(), rowNum, i) );
                cellData.setCellComment( comment.getString().getString() );
                commentDataCells.add( cellData );
            }
        }
        return commentDataCells;
    }

}
