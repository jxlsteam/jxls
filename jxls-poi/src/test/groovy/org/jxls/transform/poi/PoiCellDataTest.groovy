package org.jxls.transform.poi

import org.apache.poi.common.usermodel.HyperlinkType
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.jxls.common.CellRef
import org.jxls.common.Context
import org.jxls.transform.TransformationConfig
import org.jxls.transform.Transformer
import spock.lang.Ignore
import spock.lang.Specification
/**
 * @author Leonid Vysochyn
 * Date: 1/30/12 5:52 PM
 */
class PoiCellDataTest extends Specification{
    Workbook wb;

    def setup(){
        wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("sheet 1")
        Row row0 = sheet.createRow(0)
        row0.createCell(0).setCellValue(1.5)
        row0.createCell(1).setCellValue('${x}')
        row0.createCell(2).setCellValue('${x*y}')
        row0.createCell(3).setCellValue('$[B2+B3]')
        Row row1 = sheet.createRow(1)
        row1.createCell(1).setCellFormula("SUM(A1:A3)")
        row1.createCell(2).setCellValue('${y*y}')
        row1.createCell(3).setCellValue('${x} words')
        row1.createCell(4).setCellValue('$[${myvar}*SUM(A1:A5) + ${myvar2}]')
        row1.createCell(5).setCellValue('$[SUM(U_(B1,B2)]')
        Row row2 = sheet.createRow(2)
        row2.createCell(0).setCellValue("XYZ")
        row2.createCell(1).setCellValue('${2*y}')
        row2.createCell(2).setCellValue('${4*4}')
        row2.createCell(3).setCellValue('${2*x}x and ${2*y}y')
        row2.createCell(4).setCellValue('${2*x}x and ${2*y} ${cur}')
        Sheet sheet2 = wb.createSheet("sheet 2")
        sheet2.createRow(0).createCell(0)
        sheet2.createRow(1).createCell(1);
//        sheet2.getRow(1).createCell(2).setCellValue('''${poi.hyperlink('http://google.com/', 'URL')}''')
        sheet2.getRow(1).createCell(2).setCellValue('''${util.hyperlink('http://google.com/', 'Google', 'URL')}''')
    }

    def "test get cell Value"(){
        when:
            PoiCellData cellData = PoiCellData.createCellData(null, new CellRef("sheet 1", row, col), wb.getSheetAt(0).getRow(row).getCell(col) )
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
            PoiCellData cellData = PoiCellData.createCellData(null, new CellRef("sheet 1", 0, 1), wb.getSheetAt(0).getRow(0).getCell(1))
            def context = new Context()
            context.putVar("x", 35)
            cellData.transformer = Mock(Transformer)
            cellData.getTransformer().getTransformationConfig() >> new TransformationConfig()
        expect:
            cellData.evaluate(context) == 35
    }
    
    def "test evaluate multiple regex"(){
        setup:
            PoiCellData cellData = PoiCellData.createCellData(null, new CellRef("sheet 1", 2, 3),wb.getSheetAt(0).getRow(2).getCell(3))
            def context = new Context()
            context.putVar("x", 2)
            context.putVar("y", 3)
            cellData.transformer = Mock(Transformer)
            cellData.getTransformer().getTransformationConfig() >> new TransformationConfig()
        expect:
            cellData.evaluate(context) == "4x and 6y"
    }

    def "test evaluate single expression constant string concatenation"(){
        setup:
            PoiCellData cellData = PoiCellData.createCellData(null, new CellRef("sheet 1", 1, 3),wb.getSheetAt(0).getRow(1).getCell(3))
            def context = new Context()
            context.putVar("x", 35)
            cellData.transformer = Mock(Transformer)
            cellData.getTransformer().getTransformationConfig() >> new TransformationConfig()
        expect:
            cellData.evaluate(context) == "35 words"
    }

    def "test evaluate regex with dollar sign"(){
        PoiCellData cellData = PoiCellData.createCellData(null, new CellRef("sheet 1", 2, 4), wb.getSheetAt(0).getRow(2).getCell(4))
        def context = new Context()
        context.putVar("x", 2)
        context.putVar("y", 3)
        context.putVar("cur", '$')
        cellData.transformer = Mock(Transformer)
        cellData.getTransformer().getTransformationConfig() >> new TransformationConfig()
        expect:
            cellData.evaluate(context) == '4x and 6 $'
    }

    def "test write to another sheet"(){
        setup:
            PoiTransformer transformer = Mock(PoiTransformer)
            PoiRowData rowData = PoiRowData.createRowData(Mock(PoiSheetData), Mock(Row), transformer)
            PoiCellData cellData = PoiCellData.createCellData(rowData, new CellRef("sheet 1", 0, 1),wb.getSheetAt(0).getRow(0).getCell(1))
            def context = new Context()
            context.putVar("x", 35)
            Cell targetCell = wb.getSheetAt(1).getRow(0).getCell(0)
            cellData.transformer = Mock(Transformer)
        cellData.getTransformer().getTransformationConfig() >> new TransformationConfig()
        when:
            cellData.writeToCell(targetCell, context, null)
        then:
            wb.getSheetAt(1).getRow(0).getCell(0).getNumericCellValue() == 35
    }

    def "test write BigDecimal"(){
        setup:
            PoiTransformer transformer = Mock(PoiTransformer)
            PoiRowData rowData = PoiRowData.createRowData(Mock(PoiSheetData), Mock(Row), transformer)
            PoiCellData cellData = PoiCellData.createCellData(rowData, new CellRef("sheet 1", 0, 1), wb.getSheetAt(0).getRow(0).getCell(1))
            def context = new Context()
            BigDecimal xValue = new BigDecimal(1234.56D)
            context.putVar("x", xValue)
            cellData.transformer = Mock(Transformer)
            cellData.getTransformer().getTransformationConfig() >> new TransformationConfig()
        when:
            cellData.writeToCell(wb.getSheetAt(1).getRow(1).getCell(1), context, null)
        then:
            xValue == new BigDecimal( wb.getSheetAt(1).getRow(1).getCell(1).getNumericCellValue() )
    }

    def "test write Date"(){
        setup:
            PoiTransformer transformer = Mock(PoiTransformer)
            PoiRowData rowData = PoiRowData.createRowData(Mock(PoiSheetData), Mock(Row), transformer)
            PoiCellData cellData = PoiCellData.createCellData(rowData, new CellRef("sheet 1", 0, 1), wb.getSheetAt(0).getRow(0).getCell(1))
            def context = new Context()
            Date today = new Date()
            context.putVar("x", today)
            cellData.transformer = transformer
            cellData.getTransformer().getTransformationConfig() >> new TransformationConfig()
        when:
            cellData.writeToCell(wb.getSheetAt(1).getRow(1).getCell(1), context, null)
        then:
            today == wb.getSheetAt(1).getRow(1).getCell(1).getDateCellValue()
    }

    def "test write user formula"(){
        setup:
            PoiTransformer transformer = Mock(PoiTransformer)
            PoiRowData rowData = PoiRowData.createRowData(Mock(PoiSheetData), Mock(Row), transformer)
            PoiCellData cellData = PoiCellData.createCellData(rowData, new CellRef("sheet 1", 0, 3),wb.getSheetAt(0).getRow(0).getCell(3))
            def context = new Context()
            cellData.transformer = Mock(Transformer)
            cellData.getTransformer().getTransformationConfig() >> new TransformationConfig()
        when:
            cellData.writeToCell(wb.getSheetAt(1).getRow(1).getCell(1), context, null)
        then:
            wb.getSheetAt(1).getRow(1).getCell(1).getCellFormula() == "B2+B3"
    }
    
    def "test write parameterized formula cell"(){
        setup:
            PoiTransformer transformer = Mock(PoiTransformer)
            PoiRowData rowData = PoiRowData.createRowData(Mock(PoiSheetData), Mock(Row), transformer)
            PoiCellData cellData = PoiCellData.createCellData(rowData, new CellRef("sheet 1", 1, 4),wb.getSheetAt(0).getRow(1).getCell(4))
            def context = new Context()
            context.putVar("myvar", 2)
            context.putVar("myvar2", 3)
            wb.getSheetAt(0).createRow(7).createCell(7)
            cellData.transformer = Mock(Transformer)
            cellData.getTransformer().getTransformationConfig() >> new TransformationConfig()
        when:
            cellData.writeToCell(wb.getSheetAt(0).getRow(7).getCell(7), context, null)
        then:
            wb.getSheetAt(0).getRow(7).getCell(7).getCellFormula() == "2*SUM(A1:A5)+3"
    }
    
    def "test formula cell check"(){
        when:
            PoiTransformer transformer = Mock(PoiTransformer)
            PoiRowData rowData = PoiRowData.createRowData(Mock(PoiSheetData), Mock(Row), transformer)
            PoiCellData notFormulaCell = PoiCellData.createCellData(rowData, new CellRef("sheet 1", 0, 1), wb.getSheetAt(0).getRow(0).getCell(1))
            PoiCellData formulaCell1 = PoiCellData.createCellData(rowData, new CellRef("sheet 1", 1, 1), wb.getSheetAt(0).getRow(1).getCell(1))
            PoiCellData formulaCell2 = PoiCellData.createCellData(rowData, new CellRef("sheet 1", 0, 3), wb.getSheetAt(0).getRow(0).getCell(3))
        then:
            !notFormulaCell.isFormulaCell()
            formulaCell1.isFormulaCell()
            formulaCell2.isFormulaCell()
//            formulaCell3.isFormulaCell()
//            formulaCell3.getFormula() == "B2+B3"
    }

    def "test write formula with jointed cells"(){
        setup:
            PoiTransformer transformer = Mock(PoiTransformer)
            PoiRowData rowData = PoiRowData.createRowData(Mock(PoiSheetData), Mock(Row), transformer)
            PoiCellData cellData = PoiCellData.createCellData(rowData, new CellRef("sheet 1", 1, 5), wb.getSheetAt(0).getRow(1).getCell(5))
            def context = new Context()
            cellData.transformer = Mock(Transformer)
            cellData.getTransformer().getTransformationConfig() >> new TransformationConfig()
        when:
            cellData.writeToCell(wb.getSheetAt(1).getRow(1).getCell(1), context, null)
        then:
            wb.getSheetAt(1).getRow(1).getCell(1).getStringCellValue() == "SUM(U_(B1,B2)"
    }

    // todo:
    @Ignore("Not implemented yet")
    def "test write merged cell"(){

    }

    def "test hyperlink cell"(){
        setup:
            PoiCellData cellData = PoiCellData.createCellData(null, new CellRef("sheet 2", 1, 2), wb.getSheetAt(1).getRow(1).getCell(2))
            def poiContext = new PoiContext()
            cellData.transformer = Mock(Transformer)
            cellData.getTransformer().getTransformationConfig() >> new TransformationConfig()
        when:
            cellData.writeToCell(wb.getSheetAt(1).getRow(1).getCell(2), poiContext, null)
        then:
            def hyperlink = wb.getSheetAt(1).getRow(1).getCell(2).getHyperlink()
            hyperlink != null
            hyperlink.address == "http://google.com/"
            wb.getSheetAt(1).getRow(1).getCell(2).getStringCellValue() == "Google"
            hyperlink.typeEnum == HyperlinkType.URL
    }

}
