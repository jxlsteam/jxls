package org.jxls.transform.poi

import spock.lang.Specification
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Row

/**
 * @author Leonid Vysochyn
 * Date: 2/1/12 2:03 PM
 */
class PoiRowDataTest extends Specification{
    Workbook wb;

    def setup(){
        wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("sheet 1")
        Row row0 = sheet.createRow(0)
        row0.createCell(0).setCellValue(1.5)
        row0.createCell(1).setCellValue('${x}')
        row0.createCell(2).setCellValue('${x*y}')
        row0.setHeight((short)123)
        Row row1 = sheet.createRow(1)
        row1.createCell(1).setCellFormula("SUM(A1:A3)")
        row1.createCell(2).setCellValue('${y*y}')
        row1.createCell(3).setCellValue('${x} words')
        row1.setHeight((short)456)
        Row row2 = sheet.createRow(2)
        row2.createCell(0).setCellValue("XYZ")
        row2.createCell(1).setCellValue('${2*y}')
        row2.createCell(2).setCellValue('${4*4}')
        row2.createCell(3).setCellValue('${2*x}x and ${2*y}y')
        row2.createCell(4).setCellValue('${2*x}x and ${2*y} ${cur}')
        sheet.setColumnWidth(1, 123);
        sheet.setColumnBreak(3);
        Sheet sheet2 = wb.createSheet("sheet 2")
        sheet2.createRow(0).createCell(0)
    }

    def "test createRowData"(){
        when:
            def row1 = wb.getSheetAt(0).getRow(1)
            def rowData = PoiRowData.createRowData(Mock(PoiSheetData), row1, null);
        then:
            rowData.getHeight() == row1.getHeight()
            rowData.getNumberOfCells() == 4
            rowData.getCellData(2).getCellValue() == '${y*y}'
            rowData.getCellData(2).getSheetName() == "sheet 1"
    }
    
    def "test createRowData for null row"(){
        given:
            def row5 = wb.getSheetAt(0).getRow(5)
        expect:
            PoiRowData.createRowData(Mock(PoiSheetData), row5, null) == null
    }
}
