package com.jxls.writer.transform.poi

import spock.lang.Specification
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Row

/**
 * @author Leonid Vysochyn
 * Date: 2/1/12 12:05 PM
 */
class SheetDataTest extends Specification{
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

    def "test read sheet data"(){
        when:
            Sheet sheet = wb.getSheetAt(0)
            SheetData sheetData = SheetData.createSheetData(sheet)
        then:
            sheet.getSheetName() == sheetData.getSheetName()
            sheet.getColumnWidth(2) == sheetData.getColumnWidth(2)
            sheet.getColumnWidth(1) == sheetData.getColumnWidth(1)
            sheet.getRow(0).getHeight() == sheetData.getRowData(0).getHeight()
    }
}
