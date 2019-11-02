package org.jxls.transform.poi;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class PoiTestHelper {
    public static double cellNumericValue(Workbook workbook, int rowNum, int colNum){
        Sheet sheet = workbook.getSheetAt(0);
        return sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
    }

    public static double cellNumericValue(Workbook workbook, String cell){
        Sheet sheet = workbook.getSheetAt(0);
        CellRangeAddress cellAddress = CellRangeAddress.valueOf(cell);
        return sheet.getRow(cellAddress.getFirstRow()).getCell(cellAddress.getFirstColumn()).getNumericCellValue();
    }
}
