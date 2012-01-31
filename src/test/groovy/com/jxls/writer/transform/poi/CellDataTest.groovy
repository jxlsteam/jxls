package com.jxls.writer.transform.poi

import spock.lang.Specification
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Workbook
import com.jxls.writer.command.Context

/**
 * @author Leonid Vysochyn
 * Date: 1/30/12 5:52 PM
 */
class CellDataTest extends Specification{
    Workbook wb;

    def setup(){
        wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("sheet 1")
        Row row0 = sheet.createRow(0)
        row0.createCell(0).setCellValue(1.5)
        row0.createCell(1).setCellValue('${x}')
        row0.createCell(2).setCellValue('${x*y}')
        Row row1 = sheet.createRow(1)
        row1.createCell(1).setCellFormula("SUM(A1:A3)")
        row1.createCell(2).setCellValue('${y*y}')
        row1.createCell(3).setCellValue('${x} words')
        Row row2 = sheet.createRow(2)
        row2.createCell(0).setCellValue("XYZ")
        row2.createCell(1).setCellValue('${2*y}')
        row2.createCell(2).setCellValue('${4*4}')
        row2.createCell(3).setCellValue('${2*x}x and ${2*y}y')
        row2.createCell(4).setCellValue('${2*x}x and ${2*y} ${cur}')
    }

    def "test get cell Value"(){
        given:
            CellData cellData = null
        when:
            cellData = CellData.createCellData( wb.getSheetAt(0).getRow(row).getCell(col) )
        then:
            assert cellData.getCellValue() == value
        where:
            row | col   | value
            0   | 0     | new Double(1.5)
            0   | 1     | '${x}'
            0   | 2     | '${x*y}'
            1   | 1     | "SUM(A1:A3)"
            2   | 0     | "XYZ"
    }

    def "test evaluate simple expression"(){
        setup:
            CellData cellData = CellData.createCellData(wb.getSheetAt(0).getRow(0).getCell(1))
            def context = new Context()
            context.putVar("x", 35)
        expect:
            cellData.evaluate(context) == 35
    }
    
    def "test evaluate multiple regex"(){
        setup:
            CellData cellData = CellData.createCellData(wb.getSheetAt(0).getRow(2).getCell(3))
            def context = new Context()
            context.putVar("x", 2)
            context.putVar("y", 3)
        expect:
            cellData.evaluate(context) == "4x and 6y"
    }

    def "test evaluate single expression constant string concatenation"(){
        setup:
            CellData cellData = CellData.createCellData(wb.getSheetAt(0).getRow(1).getCell(3))
            def context = new Context()
            context.putVar("x", 35)
        expect:
            cellData.evaluate(context) == "35 words"
    }

    def "test evaluate regex with dollar sign"(){
        CellData cellData = CellData.createCellData(wb.getSheetAt(0).getRow(2).getCell(4))
        def context = new Context()
        context.putVar("x", 2)
        context.putVar("y", 3)
        context.putVar("cur", '$')
        expect:
            cellData.evaluate(context) == '4x and 6 $'
    }
}
