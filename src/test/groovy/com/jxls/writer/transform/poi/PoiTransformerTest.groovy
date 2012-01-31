package com.jxls.writer.transform.poi

import spock.lang.Specification
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Row

import com.jxls.writer.Cell
import com.jxls.writer.command.Context
import spock.lang.Ignore

/**
 * @author Leonid Vysochyn
 * Date: 1/23/12 3:23 PM
 */
class PoiTransformerTest extends Specification{
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
        Row row2 = sheet.createRow(2)
        row2.createCell(0).setCellValue("XYZ")
        row2.createCell(1).setCellValue('${2*y}')
        row2.createCell(2).setCellValue('${4*4}')
        row2.createCell(3).setCellValue('${2*x}x and ${2*y}y')
    }

    def "test template cells storage"(){
        when:
            def poiTransformer = new PoiTransformer(wb)
            wb.removeSheetAt(0)
        then:
            assert wb.getNumberOfSheets() == 0
            assert poiTransformer.getCellData(col,row,0).getCellValue() == value
        where:
            row | col   | value
            0   | 0     | new Double(1.5)
            0   | 1     | '${x}'
            0   | 2     | '${x*y}'
            1   | 1     | "SUM(A1:A3)"
            2   | 0     | "XYZ"
    }

    def "test transform string var"(){
        given:
            def poiTransformer = new PoiTransformer(wb)
            def context = new Context()
            context.putVar("x", "Abcde")
        when:
            poiTransformer.transform(new Cell(1,0), new Cell(7,7), context)
        then:
            Sheet sheet = wb.getSheetAt(0)
            Row row7 = sheet.getRow(7)
            row7.getCell(7).getStringCellValue() == "Abcde"
    }

    def "test transform numeric var"(){
        given:
            def poiTransformer = new PoiTransformer(wb)
            def context = new Context()
            context.putVar("x", 3)
            context.putVar("y", 5)
        when:
            poiTransformer.transform(new Cell(2,0), new Cell(7,7), context)
        then:
            Sheet sheet = wb.getSheetAt(0)
            Row row7 = sheet.getRow(7)
            row7.getCell(7).cellType == org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC
            row7.getCell(7).getNumericCellValue() == 15
    }

    def "test transform formula cell"(){
        given:
            def poiTransformer = new PoiTransformer(wb)
            def context = new Context()
        when:
            poiTransformer.transform(new Cell(1,1), new Cell(7,7), context)
        then:
            Sheet sheet = wb.getSheetAt(0)
            Row row7 = sheet.getRow(7)
            row7.getCell(7).cellType == org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA
            row7.getCell(7).getCellFormula() == "SUM(A1:A3)"
    }

    def "test transform a cell to other sheet"(){
        given:
            def poiTransformer = new PoiTransformer(wb)
            def context = new Context()
            context.putVar("x", "Abcde")
        when:
            poiTransformer.transform(new Cell(1,0), new Cell(7,7,1), context)
        then:
            Sheet sheet = wb.getSheetAt(0)
            Row row = sheet.getRow(7)
            row == null
            Sheet sheet1 = wb.getSheetAt(1)
            Row row1 = sheet1.getRow(7)
            row1.getCell(7).getStringCellValue() == "Abcde"
    }
    
    def "test transform multiple times"(){
        given:
            def poiTransformer = new PoiTransformer(wb)
            def context1 = new Context()
            context1.putVar("x", "Abcde")
            def context2 = new Context()
            context2.putVar("x", "Fghij")
        when:
            poiTransformer.transform(new Cell(1,0), new Cell(1, 5), context1)
            poiTransformer.transform(new Cell(1,0), new Cell(1, 10), context2)
        then:
            Sheet sheet = wb.getSheetAt(0)
            Row row5 = sheet.getRow(5)
            Row row10 = sheet.getRow(10)
            row5.getCell(1).getStringCellValue() == "Abcde"
            row10.getCell(1).getStringCellValue() == "Fghij"
    }

    def "test transform overridden cells"(){
        given:
        def poiTransformer = new PoiTransformer(wb)
        def context1 = new Context()
        context1.putVar("x", "Abcde")
        def context2 = new Context()
        context2.putVar("x", "Fghij")
        when:
        poiTransformer.transform(new Cell(1,0), new Cell(1, 5), context1)
        poiTransformer.transform(new Cell(0,0), new Cell(1, 0), context1)
        poiTransformer.transform(new Cell(1,0), new Cell(1, 10), context2)
        then:
        Sheet sheet = wb.getSheetAt(0)
        Row row5 = sheet.getRow(5)
        Row row10 = sheet.getRow(10)
        row5.getCell(1).getStringCellValue() == "Abcde"
        row10.getCell(1).getCellType() == org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING
        row10.getCell(1).getStringCellValue() == "Fghij"
    }

    def "test multiple expressions in a cell"(){
        given:
            def poiTransformer = new PoiTransformer(wb)
            def context = new Context()
            context.putVar("x", 2)
            context.putVar("y", 3)
        when:
            poiTransformer.transform(new Cell(3,2), new Cell(7,7), context)
        then:
            Sheet sheet = wb.getSheetAt(0)
            Row row7 = sheet.getRow(7)
            row7.getCell(7).getStringCellValue() == "4x and 6y"
    }

}
