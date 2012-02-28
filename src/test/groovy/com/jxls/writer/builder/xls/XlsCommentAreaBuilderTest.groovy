package com.jxls.writer.builder.xls

import spock.lang.Specification
import spock.lang.FailsWith
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.usermodel.Cell
import com.jxls.writer.area.Area
import org.apache.poi.ss.usermodel.Drawing
import org.apache.poi.ss.usermodel.CreationHelper
import org.apache.poi.ss.usermodel.ClientAnchor
import org.apache.poi.ss.usermodel.Comment
import org.apache.poi.ss.usermodel.RichTextString
import spock.lang.Ignore
import com.jxls.writer.transform.Transformer
import com.jxls.writer.transform.poi.PoiTransformer
import com.jxls.writer.util.Util
import com.jxls.writer.transform.poi.PoiUtil
import com.jxls.writer.common.CellData
import com.jxls.writer.common.CellRef

/**
 * @author Leonid Vysochyn
 */
class XlsCommentAreaBuilderTest extends Specification {
    Workbook wb;
    
    def setup(){
        wb = new HSSFWorkbook()
        Sheet sheet = wb.createSheet("sheet1")
        Row row0 = sheet.createRow(0)
        row0.createCell(0).setCellValue(1.5)
        Row row1 = sheet.createRow(1)
        setCellComment(row1.createCell(1),"jx:each(items='departments', var='department', lastCell='F13')")
        setCellComment(row1.createCell(2), "each(items='items', var='item', lastCell='E2')")
        Row row2 = sheet.createRow(2)
        row2.createCell(0).setCellValue("XYZ")
        row2.createCell(1).setCellValue('${2*x}')
        Row row4 = sheet.createRow(4)
        row4.createCell(0).setCellValue('${department.chief.name}')
        Row row8 = sheet.createRow(8)
        Cell cell = row8.createCell(0)
        cell.setCellValue('${employee.name}')
        setCellComment(cell, """
        jx:each(items='department.staff' var="employee"      lastCell="F9")
        jx:if(condition="employee.payment <= 2000", lastCell="F9",  areas=["A9:F9","A18:F18"])""")
        sheet.createRow(9).createCell(0).setCellValue("Totals")
        sheet.createRow(11).createCell(5).setCellValue('$[F5+F10]')
        sheet.createRow(17).createCell(0).setCellValue('${employee.name}')
        Sheet sheet2 = wb.createSheet("sheet2")
        setCellComment(sheet2.createRow(1).createCell(1), 'jx:each(items="myItems"    var="mywar"      lastCell="D2" areas=["C9:F9"])')
        setCellComment(sheet2.createRow(1).createCell(7), 'jx:if(condition="myvar.value==2" lastCell="K2" )')
        setCellComment(sheet2.createRow(4).createCell(0), ' jx:each( items = "employees" var="employee" lastCell="D5") ')
    }
    @Ignore
    def "integration test for build"(){
        def transformer = PoiTransformer.createTransformer(wb);
        when:
            def areaBuilder = new XlsCommentAreaBuilder(transformer)
            List<Area> areas = areaBuilder.build()
        then:
            areas.size() == 3
    }

    def "unit test for build"(){
            def transformer = Mock(Transformer)
            def cellData0 = new CellData(new CellRef("sheet1!B1"))
            cellData0.setCellComment("jx:area(lastCell='G8')")
            def cellData1 = new CellData(new CellRef("sheet1!B2"))
            cellData1.setCellComment("jx:each(items='departments', var='department', lastCell='F13')")
            def cellData2 = new CellData(new CellRef("sheet1!C2"))
            cellData2.setCellComment("each(items='items', var='item', lastCell='E2')")
            def cellData3 = new CellData(new CellRef("sheet1!A9"))
            cellData3.setCellComment("""
                jx:each(items='department.staff' var="employee"      lastCell="F9")
                jx:if(condition="employee.payment <= 2000", lastCell="F9",  areas=["A9:F9","A18:F18"])""" )
            def cellData4 = new CellData(new CellRef("sheet2!B2"))
            cellData4.setCellComment('''
                    jx:area(lastCell = "K10")
                    jx:each(items="myItems"    var="mywar"      lastCell="D2" areas=["C9:F9"])''')
            def cellData5 = new CellData(new CellRef("sheet2", 1, 7))
            cellData5.setCellComment('jx:if(condition="myvar.value==2" lastCell="K2" )')
            def cellData6 = new CellData(new CellRef("sheet2", 4, 0))
            cellData6.setCellComment(' jx:each( items = "employees" var="employee" lastCell="D5") ')
        when:
            def areaBuilder = new XlsCommentAreaBuilder(transformer)
            List<Area> areas = areaBuilder.build()
        then:
            transformer.getCommentedCells() >> [cellData1, cellData2, cellData3, cellData4, cellData5, cellData6]
            areas.size() == 2

    }

    static def setCellComment(Cell cell, String commentText){
        PoiUtil.setCellComment(cell, commentText, "jxlswriter", null)
    }
    
}
