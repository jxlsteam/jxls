package com.jxls.writer.transform.poi

import spock.lang.Specification
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Cell
import com.jxls.writer.Pos
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
        row0.createCell(0).setCellValue(1)
        row0.createCell(1).setCellValue('${x}')
        row0.createCell(2).setCellValue('${x*y}')
        Row row1 = sheet.createRow(1)
        row1.createCell(1).setCellValue("Abc")
        row1.createCell(2).setCellValue('${y*y}')
        Row row2 = sheet.createRow(2)
        row2.createCell(0).setCellValue("XYZ")
        row2.createCell(1).setCellValue('${2*y}')
        row2.createCell(2).setCellValue('${4*4}')
    }


    def "test transform"(){
        given:
            def poiTransformer = new PoiTransformer(wb)
            def context = new Context()
            context.putVar("x", 3)
            context.putVar("y", 5)
        when:
            poiTransformer.transform(new Pos(0,0), new Pos(4,4), context)
            poiTransformer.transform(new Pos(1,0), new Pos(5,4), context)
            poiTransformer.transform(new Pos(2,0), new Pos(6,4), context)
            poiTransformer.transform(new Pos(0,1), new Pos(4,5), context)
            poiTransformer.transform(new Pos(1,1), new Pos(5,5), context)
            poiTransformer.transform(new Pos(2,1), new Pos(6,5), context)
            poiTransformer.transform(new Pos(0,2), new Pos(4,6), context)
            poiTransformer.transform(new Pos(1,2), new Pos(5,6), context)
            poiTransformer.transform(new Pos(2,2), new Pos(6,6), context)
        then:
            Sheet sheet = wb.getSheetAt(0)
            Row row0 = sheet.getRow(4)
            Row row1 = sheet.getRow(5)
            //Row row2 = sheet.getRow(6)
            assert Math.abs(row0.getCell(4).getNumericCellValue() - 1) < 1e-6
            assert Math.abs(row0.getCell(5).getNumericCellValue() - 3) < 1e-6
            assert Math.abs(row0.getCell(6).getNumericCellValue() - 3*5) < 1e-6
            assert row1.getCell(4) == null
            assert row1.getCell(5).getStringCellValue() == "Abc"
    }
}
