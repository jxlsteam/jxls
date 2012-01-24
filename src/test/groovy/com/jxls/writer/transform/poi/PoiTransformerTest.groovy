package com.jxls.writer.transform.poi

import spock.lang.Specification
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Row

import com.jxls.writer.Cell
import com.jxls.writer.command.Context

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
            assert row7.getCell(7).getStringCellValue() == "Abcde"
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
            assert row7.getCell(7).getNumericCellValue() == 15
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
            assert row7.getCell(7).cellType == org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA
            assert row7.getCell(7).getCellFormula() == "SUM(A1:A3)"
    }

}
