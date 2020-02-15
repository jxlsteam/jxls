package org.jxls.transform.poi

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.PrintSetup
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.util.IOUtils
import org.jxls.common.AreaRef
import org.jxls.common.CellData
import org.jxls.common.CellRef
import org.jxls.common.Context
import org.jxls.common.ImageType
import spock.lang.Ignore
import spock.lang.Specification

/**
 * @author Leonid Vysochyn
 * Date: 1/23/12 3:23 PM
 */
class PoiTransformerTest extends Specification{
    Workbook wb;
    CellStyle customStyle;
    byte[] workbookBytes;

    def setup(){
        wb = new HSSFWorkbook();
        customStyle = wb.createCellStyle();
        customStyle.setFillBackgroundColor(IndexedColors.AQUA.getIndex())
        customStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex())
        Font font = wb.createFont();
        font.setFontHeightInPoints((short)24);
        font.setFontName("Courier New");
        font.setItalic(true);
        customStyle.setFont( font );
        Sheet sheet = wb.createSheet("sheet 1")
        PrintSetup printSetup = sheet.getPrintSetup()
        printSetup.paperSize = PrintSetup.A3_PAPERSIZE
        printSetup.footerMargin = 15
        Row row0 = sheet.createRow(0)
        row0.createCell(0).setCellValue(1.5)
        row0.createCell(1).setCellValue('${x}')
        row0.getCell(1).setCellStyle( customStyle )
        row0.createCell(2).setCellValue('${x*y}')
        PoiUtil.setCellComment( row0.getCell(2), "comment 1", "leo", null )
        row0.createCell(3).setCellValue('Merged value')
        sheet.addMergedRegion(new CellRangeAddress(0,1,3,4))
        row0.setHeight((short)23)
        sheet.setColumnWidth(1, 123);
        Row row1 = sheet.createRow(1)
        row1.createCell(1).setCellFormula("SUM(A1:A3)")
        PoiUtil.setCellComment( row1.getCell(1), "comment 2", "leo", null )
        row1.createCell(2).setCellValue('${y*y}')
        row1.setHeight((short)456)
        Row row2 = sheet.createRow(2)
        row2.createCell(0).setCellValue("XYZ")
        row2.createCell(1).setCellValue('${2*y}')
        row2.createCell(2).setCellValue('${4*4}')
        row2.createCell(3).setCellValue('${2*x}x and ${2*y}y')
        row2.createCell(4).setCellValue('$[${myvar}*SUM(A1:A5) + ${myvar2}]')
        def cell5 = row2.createCell(5)
        PoiUtil.setCellComment( cell5, "Test comment", "leo", null );
        row2.removeCell(cell5)
        ByteArrayOutputStream os = new ByteArrayOutputStream()
        wb.write(os)
        workbookBytes = os.toByteArray()
    }

    def "test initial context creation"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
        when:
            Context context = poiTransformer.createInitialContext()
        then:
            def utilClass = context.getVar(PoiTransformer.POI_CONTEXT_KEY)
            assert utilClass != null
            assert utilClass instanceof PoiUtil
    }

    def "test template cells storage"(){
        when:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            wb.removeSheetAt(0)
        then:
            assert wb.getNumberOfSheets() == 0
            CellData cellData = poiTransformer.getCellData(new CellRef("sheet 1", 0, 1))
            cellData instanceof PoiCellData
            ((PoiCellData)cellData).getCellStyle() == customStyle
            assert poiTransformer.getCellData(new CellRef("sheet 1", row, col)).getCellValue() == value
        where:
            row | col   | value
            0   | 0     | new Double(1.5)
            0   | 1     | '${x}'
            0   | 2     | '${x*y}'
            1   | 1     | "SUM(A1:A3)"
            2   | 0     | "XYZ"
            2   | 4     |  '$[${myvar}*SUM(A1:A5) + ${myvar2}]'
    }

    def "test transform string var"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
            context.putVar("x", "Abcde")
        when:
            poiTransformer.transform(new CellRef("sheet 1", 0, 1), new CellRef("sheet 1", 7, 7), context, true)
        then:
            Sheet sheet = wb.getSheetAt(0)
            Row row7 = sheet.getRow(7)
            row7.getCell(7).getStringCellValue() == "Abcde"
            sheet.getColumnWidth(7) == 123
            sheet.getRow(7).getHeight() == 23
            sheet.getRow(7).getCell(7).getCellStyle() == customStyle
    }

    def "test transform numeric var"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
            context.putVar("x", 3)
            context.putVar("y", 5)
        when:
            poiTransformer.transform(new CellRef("sheet 1", 0, 2), new CellRef("sheet 2", 7, 7), context, false)
        then:
            Sheet sheet = wb.getSheet("sheet 2")
            Row row7 = sheet.getRow(7)
            row7.getCell(7).cellTypeEnum == CellType.NUMERIC
            row7.getCell(7).getNumericCellValue() == 15
    }

    def "test transform formula cell"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
        when:
            poiTransformer.transform(new CellRef("sheet 1", 1, 1), new CellRef("sheet 2", 7, 7), context, false)
        then:
            Sheet sheet = wb.getSheet("sheet 2")
            Row row7 = sheet.getRow(7)
            row7.getCell(7).cellTypeEnum == CellType.FORMULA
            row7.getCell(7).getCellFormula() == "SUM(A1:A3)"
    }

    def "test transform a cell to other sheet"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
            context.putVar("x", "Abcde")
        when:
            poiTransformer.transform(new CellRef("sheet 1", 0, 1), new CellRef("sheet2", 7, 7), context, false)
        then:
            Sheet sheet = wb.getSheet("sheet 1")
            Row row = sheet.getRow(7)
            row == null
            Sheet sheet1 = wb.getSheet("sheet2")
            Row row1 = sheet1.getRow(7)
            row1.getCell(7).getStringCellValue() == "Abcde"
            sheet1.printSetup.footerMargin == 15
            sheet1.printSetup.paperSize == PrintSetup.A3_PAPERSIZE

    }
    
    def "test transform multiple times"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context1 = new Context()
            context1.putVar("x", "Abcde")
            def context2 = new Context()
            context2.putVar("x", "Fghij")
        when:
            poiTransformer.transform(new CellRef("sheet 1", 0, 1), new CellRef("sheet 1", 5, 1), context1, false)
            poiTransformer.transform(new CellRef("sheet 1", 0, 1), new CellRef("sheet 2", 10, 1), context2, false)
        then:
            Sheet sheet = wb.getSheet("sheet 1")
            Sheet sheet2 = wb.getSheet("sheet 2")
            Row row5 = sheet.getRow(5)
            Row row10 = sheet2.getRow(10)
            row5.getCell(1).getStringCellValue() == "Abcde"
            row10.getCell(1).getStringCellValue() == "Fghij"
    }

    def "test transform overridden cells"(){
        given:
        def poiTransformer = PoiTransformer.createTransformer(wb)
        def context1 = new Context()
        context1.putVar("x", "Abcde")
        def context2 = new Context()
        context2.putVar("x", "Fghij")
        when:
        poiTransformer.transform(new CellRef("sheet 1", 0, 1), new CellRef("sheet 1", 5, 1), context1, true)
        poiTransformer.transform(new CellRef("sheet 1", 0, 0), new CellRef("sheet 2", 0, 1), context1, true)
        poiTransformer.transform(new CellRef("sheet 1", 0, 1), new CellRef("sheet 2", 10, 1), context2, true)
        then:
        Sheet sheet = wb.getSheet("sheet 1")
        Row row5 = sheet.getRow(5)
        row5.getCell(1).getStringCellValue() == "Abcde"
        Sheet sheet2 = wb.getSheet("sheet 2")
        sheet2.getRow(0).getCell(1).getNumericCellValue() == 1.5
        Row row10 = sheet2.getRow(10)
        row10.getCell(1).getCellTypeEnum() == CellType.STRING
        row10.getCell(1).getStringCellValue() == "Fghij"
    }

    def "test multiple expressions in a cell"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
            context.putVar("x", 2)
            context.putVar("y", 3)
        when:
            poiTransformer.transform(new CellRef("sheet 1", 2, 3), new CellRef("sheet 2", 7, 7), context, true)
        then:
            Sheet sheet = wb.getSheet("sheet 2")
            Row row7 = sheet.getRow(7)
            row7.getCell(7).getStringCellValue() == "4x and 6y"
    }

    def "test ignore source column and row props"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            poiTransformer.setIgnoreColumnProps(true)
            poiTransformer.setIgnoreRowProps(true)
            def context = new Context()
            context.putVar("x", "Abcde")
        when:
            poiTransformer.transform(new CellRef("sheet 1", 0, 1), new CellRef("sheet 2", 7, 7), context, true)
        then:
            Sheet sheet1 = wb.getSheet("sheet 1")
            Sheet sheet2 = wb.getSheet("sheet 2")
            sheet2.getColumnWidth(7) != sheet1.getColumnWidth(1)
            sheet1.getRow(0).getHeight() != sheet2.getRow(7).getHeight()
    }

    def "test set formula value"(){
        given:
            def poiTransformer = PoiTransformer.createTransformer(wb)
        when:
            poiTransformer.setFormula(new CellRef("sheet 2",1, 1), "SUM(B1:B5)")
        then:
            wb.getSheet("sheet 2").getRow(1).getCell(1).getCellFormula() == "SUM(B1:B5)"
    }

    def "test get formula cells"(){
        when:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
            context.putVar("myvar", 2)
            context.putVar("myvar2", 4)
            poiTransformer.transform(new CellRef("sheet 1", 2, 4), new CellRef("sheet 2", 10, 10), context, true)
            def formulaCells = poiTransformer.getFormulaCells()
        then:
            formulaCells.size() == 2
        formulaCells.contains(new CellData("sheet 1",1,1, CellData.CellType.FORMULA, "SUM(A1:A3)"))
            formulaCells.contains(new CellData("sheet 1",2,4, CellData.CellType.STRING, '$[${myvar}*SUM(A1:A5) + ${myvar2}]'))
    }
    @Ignore("The test is not used because the target cell setting was moved from Transformer to XlsArea")
    def "test get target cells"(){
        when:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
            poiTransformer.transform(new CellRef("sheet 1", 1, 1), new CellRef("sheet 2", 10, 10), context, true)
            poiTransformer.transform(new CellRef("sheet 1", 1, 1), new CellRef("sheet 1", 10, 12), context, true)
            poiTransformer.transform(new CellRef("sheet 1", 1, 1), new CellRef("sheet 1", 10, 14), context, true)
            poiTransformer.transform(new CellRef("sheet 1", 2, 1), new CellRef("sheet 2", 20, 11), context, true)
            poiTransformer.transform(new CellRef("sheet 1", 2, 1), new CellRef("sheet 1", 20, 12), context, true)
        then:
            poiTransformer.getTargetCellRef(new CellRef("sheet 1",1,1)).toArray() == [new CellRef("sheet 2",10,10), new CellRef("sheet 1",10,12), new CellRef("sheet 1",10,14)]
            poiTransformer.getTargetCellRef(new CellRef("sheet 1",2,1)).toArray() == [new CellRef("sheet 2",20,11), new CellRef("sheet 1",20,12)]
    }
    @Ignore("The test is not used because the target cell setting was moved from Transformer to XlsArea")
    def "test reset target cells"(){
        when:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
            poiTransformer.transform(new CellRef("sheet 1", 1, 1), new CellRef("sheet 1", 10, 10), context, true)
            poiTransformer.transform(new CellRef("sheet 1", 2, 1), new CellRef("sheet 1", 20, 11), context, false)
            poiTransformer.resetTargetCellRefs()
            poiTransformer.transform(new CellRef("sheet 1", 1, 1), new CellRef("sheet 2", 10, 12), context, false)
            poiTransformer.transform(new CellRef("sheet 1", 1, 1), new CellRef("sheet 1", 10, 14), context, false)
            poiTransformer.transform(new CellRef("sheet 1", 2, 1), new CellRef("sheet 1", 20, 12), context, false)
        then:
            poiTransformer.getTargetCellRef(new CellRef("sheet 1",1,1)).toArray() == [new CellRef("sheet 2",10,12), new CellRef("sheet 1",10,14)]
            poiTransformer.getTargetCellRef(new CellRef("sheet 1",2,1)).toArray() == [new CellRef("sheet 1",20,12)]
    }

    def "test transform merged cells"(){
        when:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            def context = new Context()
            wb.getSheetAt(0).getNumMergedRegions() == 1
            poiTransformer.transform(new CellRef("sheet 1", 0, 3), new CellRef("sheet 1", 10, 10), context, false)
        then:
            Sheet sheet = wb.getSheet("sheet 1")
            sheet.getNumMergedRegions() == 2
            sheet.getMergedRegion(1).toString() == new CellRangeAddress(10,11,10,11).toString()
    }

    def "test clear cell"(){
        when:
            def poiTransformer = PoiTransformer.createTransformer(wb)
            poiTransformer.clearCell(new CellRef("'sheet 1'!B1"))
        then:
            def cell = wb.getSheetAt(0).getRow(0).getCell(1)
            cell.cellTypeEnum == CellType.BLANK
            cell.stringCellValue == ""
            cell.cellStyle != customStyle
            cell.cellStyle == wb.getCellStyleAt((short)0)
    }

    def "test get commented cells"(){
        when:
            def transformer = PoiTransformer.createTransformer(wb)
            def commentedCells = transformer.getCommentedCells()
        then:
            commentedCells.size() == 3
            commentedCells.get(0).getCellComment() == "comment 1"
            commentedCells.get(1).getCellComment() == "comment 2"
            commentedCells.get(2).getCellComment() == "Test comment"
    }

    def "test addImage"(){
        given:
            InputStream imageInputStream = PoiTransformerTest.class.getResourceAsStream("/org/jxls/templatebasedtests/ja.png");
            byte[] imageBytes = IOUtils.toByteArray(imageInputStream);
        when:
            def transformer = PoiTransformer.createTransformer(wb)
            transformer.addImage(new AreaRef("'sheet 1'!A1:C10"), imageBytes, ImageType.PNG);
        then:
            def pictures = wb.getAllPictures()
            pictures.size() == 1
    }

    def "test write without output stream"(){
        given:
        InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
        def poiTransformer = PoiTransformer.createTransformer(inputStream)
        when:
        poiTransformer.write()
        then:
        def e = thrown(IllegalStateException)
        e.cause == null
        e.message == "Cannot write a workbook with an uninitialized output stream"
    }

    def "test write workbook"(){
        given:
        InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream()
        def poiTransformer = PoiTransformer.createTransformer(inputStream, outputStream)
        when:
        poiTransformer.write()
        then:
        outputStream.size() > 0
    }


}
