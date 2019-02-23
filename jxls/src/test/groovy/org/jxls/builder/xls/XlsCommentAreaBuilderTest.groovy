package org.jxls.builder.xls

import org.jxls.command.GridCommand
import spock.lang.Specification

import org.jxls.area.Area

import org.jxls.transform.Transformer
import org.jxls.common.CellData
import org.jxls.common.CellRef
import org.jxls.common.AreaRef
import org.jxls.command.EachCommand
import org.jxls.command.IfCommand
import org.jxls.area.XlsArea

/**
 * @author Leonid Vysochyn
 */
class XlsCommentAreaBuilderTest extends Specification {

    def "test build"(){
            def transformer = Mock(Transformer)
            def cellData0 = new CellData(new CellRef("sheet1!B1"))
            cellData0.setCellComment("jx:area(lastCell='G8')")
            def cellData1 = new CellData(new CellRef("sheet1!B2"))
            cellData1.setCellComment("jx:each(items='departments', var='department', lastCell='G7')")
            def cellData2 = new CellData(new CellRef("sheet1!C3"))
            cellData2.setCellComment("each(items='items', var='item', lastCell='E3')")
            def cellData3 = new CellData(new CellRef("sheet1!B7"))
            cellData3.setCellComment("""
                jx:each(items='department.staff' var="employee"      lastCell="G7")
                jx:if(condition="employee.payment <= 2000", lastCell="G7",  areas=[B7:G7,A15:F15])""" )
            def cellData4 = new CellData(new CellRef("sheet2!B2"))
            cellData4.setCellComment('''
                    jx:area(lastCell = "K10")
                    jx:each(items="myItems"    var="mywar"      lastCell="E2" areas=[ C12:F12 ])''')
            def cellData5 = new CellData(new CellRef("sheet2!G2"))
            cellData5.setCellComment('jx:if(condition="myvar.value==2" lastCell="K2" )')
            def cellData6 = new CellData(new CellRef("sheet2!C5"))
            cellData6.setCellComment(' jx:each( items = "employees" var="employee" lastCell="E5") ')
            def cellData7 = new CellData(new CellRef("sheet3!B2"))
            cellData7.setCellComment('jx:grid(headers="hsData" data="mainData")')
        when:
            def areaBuilder = new XlsCommentAreaBuilder(transformer)
            List<Area> areas = areaBuilder.build()
        then:
            transformer.getCommentedCells() >> [cellData0, cellData1, cellData2, cellData3, cellData4, cellData5, cellData6, cellData7]
            areas.size() == 2
            // area at sheet1 checks
            def area1 = areas[0]
            area1.getAreaRef() == new AreaRef("sheet1!B1:G8")
            area1 instanceof  XlsArea
            def commandDataList = area1.getCommandDataList()
            commandDataList.size() == 1
            commandDataList[0].getAreaRef() == new AreaRef("sheet1!B2:G7")
            def command1 = commandDataList[0].getCommand()
            command1.name == "each"
            ((EachCommand)command1).items == "departments"
            ((EachCommand)command1).var == "department"
            command1.getAreaList().size() == 1
            def eachArea = command1.getAreaList()[0]
            eachArea.getAreaRef() == new AreaRef("sheet1!B2:G7")
            eachArea.getCommandDataList().size() == 1
            eachArea.getCommandDataList()[0].areaRef == new AreaRef("sheet1!B7:G7")
            def command2 = eachArea.getCommandDataList()[0].getCommand()
            command2.name == "each"
            def eachArea2 = command2.getAreaList()[0]
            eachArea2.areaRef == new AreaRef("sheet1!B7:G7")
            eachArea2.getCommandDataList().size() == 1
            eachArea2.getCommandDataList()[0].areaRef == new AreaRef("sheet1!B7:G7")
            def command3 = eachArea2.getCommandDataList()[0].getCommand()
            command3.name == "if"
            command3.getAreaList().size() == 2
            def ifArea1 = command3.getAreaList()[0]
            ifArea1.areaRef == new AreaRef("sheet1!B7:G7")
            ifArea1.getCommandDataList().size() == 0
            def ifArea2 = command3.getAreaList()[1]
            ifArea2.areaRef == new AreaRef("sheet1!A15:F15")
            ifArea2.getCommandDataList().size() == 0
            // area at sheet2 checks
            def area2 = areas[1]
            area2.getAreaRef() == new AreaRef("sheet2!B2:K10")
            area2.getCommandDataList().size() == 3
            area2.getCommandDataList()[0].getAreaRef() == new AreaRef("sheet2!B2:E2")
            def command4 = area2.getCommandDataList()[0].getCommand()
            command4.name == "each"
            def eachArea3 = command4.getAreaList()[0]
            eachArea3.areaRef == new AreaRef("sheet2!C12:F12")
            eachArea3.getCommandDataList().isEmpty()
            area2.getCommandDataList()[1].areaRef == new AreaRef("sheet2!G2:K2")
            area2.getCommandDataList()[2].areaRef == new AreaRef("sheet2!C5:E5")
            def command5 = area2.getCommandDataList()[1].getCommand()
            command5.name == "if"
            command5.getAreaList().size() == 1
            command5.getAreaList()[0].areaRef == new AreaRef("sheet2!G2:K2")
            command5.getAreaList()[0].commandDataList.isEmpty()
            ((IfCommand)command5).condition == "myvar.value==2"
            def command6 = area2.getCommandDataList()[2].getCommand()
            command6.name == "each"
            command6.areaList[0].areaRef == new AreaRef("sheet2!C5:E5")
            command6.areaList[0].commandDataList.isEmpty()
            ((EachCommand)command6).items == "employees"

//            def command7 = area3.getCommandDataList()[0].getCommand()
    }

    def "test build with GridCommand"(){
        def transformer = Mock(Transformer)
        def cellData0 = new CellData(new CellRef("sheet1!B2"))
        cellData0.setCellComment("jx:area(lastCell='C4')")
        def cellData1 = new CellData(new CellRef("sheet1!B3"))
        cellData1.setCellComment("jx:grid(headers='headers', data='data', lastCell='B4', areas=[B3:B3, B4:B4])")
        when:
        def areaBuilder = new XlsCommentAreaBuilder(transformer)
        List<Area> areas = areaBuilder.build()
        then:
        transformer.getCommentedCells() >> [cellData0, cellData1]

        areas.size() == 1
        // area at sheet1 checks
        def area1 = areas[0]
        area1.getAreaRef() == new AreaRef("sheet1!B2:C4")
        area1 instanceof  XlsArea
        def commandDataList = area1.getCommandDataList()
        commandDataList.size() == 1
        commandDataList[0].getAreaRef() == new AreaRef("sheet1!B3:B4")
        def command1 = commandDataList[0].getCommand()
        command1.name == "grid"
        ((GridCommand)command1).headers == "headers"
        ((GridCommand)command1).data == "data"
        command1.getAreaList().size() == 2
        def headerArea = command1.getAreaList()[0]
        headerArea.getAreaRef() == new AreaRef("sheet1!B3:B3")
        headerArea.getCommandDataList().size() == 0
        def bodyArea = command1.getAreaList()[1]
        bodyArea.getAreaRef()  == new AreaRef("sheet1!B4:B4")
    }


}
