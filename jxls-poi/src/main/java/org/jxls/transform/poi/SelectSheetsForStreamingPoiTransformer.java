package org.jxls.transform.poi;

import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jxls.common.CellRef;

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
    protected Row getRow(CellRef srcCellRef, CellRef targetCellRef, Sheet sheet) {
        boolean useStreamingForThisSheet = useStreaming(targetCellRef.getSheetName());
        if (!useStreamingForThisSheet && isStreaming()) {
            // use "fat" data sheet for transformation
            sheet = ((SXSSFWorkbook) getWorkbook()).getXSSFWorkbook().getSheet(targetCellRef.getSheetName());
        }
        Row row = sheet.getRow(targetCellRef.getRow());
        if (row == null) {
            if (useStreamingForThisSheet && isStreaming()) { 
                XSSFSheet _sh = ((SXSSFWorkbook) getWorkbook()).getXSSFWorkbook().getSheet(targetCellRef.getSheetName());
                if (_sh.getPhysicalNumberOfRows() > 0 && targetCellRef.getRow() <= _sh.getLastRowNum()) {
                    row = _sh.getRow(targetCellRef.getRow());
                    sheet = _sh;
                }
            }
            if (row == null) { // yes, check again for null!
                row = sheet.createRow(targetCellRef.getRow());
            }
        }
        return row;
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
