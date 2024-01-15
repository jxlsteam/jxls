package org.jxls.transform.poi;

import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.Context;

/**
 * You can use this PoiTransformer implementation to decide which worksheets use streaming.
 */
public class SelectSheetsForStreamingPoiTransformer extends PoiTransformer {
	protected Set<String> dataSheetsToUseStreaming = null;
    protected boolean allSheets = false;
    
    public SelectSheetsForStreamingPoiTransformer(Workbook workbook) {
        super(workbook, true);
    }

    public SelectSheetsForStreamingPoiTransformer(Workbook workbook, boolean allSheets,
            int rowAccessWindowSize, boolean compressTmpFiles, boolean useSharedStringsTable) {
        super(workbook, true, rowAccessWindowSize, compressTmpFiles, useSharedStringsTable);
        this.allSheets = allSheets;
    }

    public SelectSheetsForStreamingPoiTransformer(Workbook workbook, Set<String> sheetNames,
            int rowAccessWindowSize, boolean compressTmpFiles, boolean useSharedStringsTable) {
        super(workbook, true, rowAccessWindowSize, compressTmpFiles, useSharedStringsTable);
        this.dataSheetsToUseStreaming = sheetNames;
    }

    public void setDataSheetsToUseStreaming(Set<String> sheetNames) {
        this.dataSheetsToUseStreaming = sheetNames;
    }

    @Override
    public void transform(CellRef srcCellRef, CellRef targetCellRef, Context context, boolean updateRowHeightFlag) {
        CellData cellData = isTransformable(srcCellRef, targetCellRef);
        if (cellData == null) {
            return;
        }
        Sheet destSheet = getWorkbook().getSheet(targetCellRef.getSheetName());
        if (destSheet == null) {
            destSheet = getWorkbook().createSheet(targetCellRef.getSheetName());
            PoiUtil.copySheetProperties(getWorkbook().getSheet(srcCellRef.getSheetName()), destSheet);
        }
        boolean useStreamingForThisSheet = useStreaming(targetCellRef.getSheetName());
        if (!useStreamingForThisSheet && isStreaming()) {
            // use "fat" data sheet for transformation
            destSheet = ((SXSSFWorkbook) getWorkbook()).getXSSFWorkbook().getSheet(targetCellRef.getSheetName());
        }
        Row destRow = destSheet.getRow(targetCellRef.getRow());
        if (destRow == null) {
            if (useStreamingForThisSheet && isStreaming()) { 
                XSSFSheet _sh = ((SXSSFWorkbook) getWorkbook()).getXSSFWorkbook().getSheet(targetCellRef.getSheetName());
                if (_sh.getPhysicalNumberOfRows() > 0 && targetCellRef.getRow() <= _sh.getLastRowNum()) {
                    destRow = _sh.getRow(targetCellRef.getRow());
                    destSheet = _sh;
                    if (destRow == null) {
                        destRow = destSheet.createRow(targetCellRef.getRow());
                    }
                } else {
                    destRow = destSheet.createRow(targetCellRef.getRow());
                }
            } else {
                destRow = destSheet.createRow(targetCellRef.getRow());
            }
        }
        transformCell(srcCellRef, targetCellRef, context, updateRowHeightFlag, cellData, destSheet, destRow);
    }
    
    protected boolean useStreaming(String sheetName) {
    	return allSheets || (dataSheetsToUseStreaming != null && dataSheetsToUseStreaming.contains(sheetName));
    }
    
    @Override
    protected Row getRowForClearCell(Sheet sheet, CellRef cellRef) {
        Row row = super.getRowForClearCell(sheet, cellRef);
        // remove comments:
        if (row == null && isStreaming()) {
            XSSFSheet _sh = getXSSFWorkbook().getSheet(cellRef.getSheetName());
            if (_sh.getPhysicalNumberOfRows() > 0 && cellRef.getRow() <= _sh.getLastRowNum()) {
                row = _sh.getRow(cellRef.getRow());
            }
        }
        return row;
    }
}
