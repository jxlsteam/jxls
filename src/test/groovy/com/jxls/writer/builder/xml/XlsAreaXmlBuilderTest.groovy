package com.jxls.writer.builder.xml

import spock.lang.Specification
import com.jxls.writer.CellRef
import com.jxls.writer.Size
import com.jxls.writer.command.EachCommand
import com.jxls.writer.command.IfCommand
import com.jxls.writer.transform.Transformer

/**
 * @author Leonid Vysochyn
 * Date: 2/14/12 11:55 AM
 */
class XlsAreaXmlBuilderTest extends Specification{
    def "test build"(){
        InputStream is = XlsAreaXmlBuilderTest.class.getResourceAsStream("xlsarea.xml")
        Transformer transformer = Mock(Transformer)
        assert is != null
        when:
            def xlsAreaList = new XlsAreaXmlBuilder(transformer).build(is)
        then:
            xlsAreaList.size() == 2
            def xlsArea = xlsAreaList.get(0)
            xlsArea != null
            xlsArea.getStartCellRef() == new CellRef("Template!A1")
            xlsArea.getInitialSize() == new Size(7,15)
            xlsArea.getTransformer() == transformer
            def commandDataList1 = xlsArea.getCommandDataList()
            commandDataList1.size() == 1
            commandDataList1.get(0).startCellRef == new CellRef("Template!A2")
            commandDataList1.get(0).size == new Size(6,11)
            def command1 = commandDataList1.get(0).getCommand()
            command1.getName() == "each"
            command1 instanceof EachCommand
            ((EachCommand)command1).items == "departments"
            ((EachCommand)command1).var == "department"
            command1.getAreaList().size() == 1
            def eachArea = command1.getAreaList().get(0)
            eachArea.getStartCellRef() == new CellRef("Template!A2")
            eachArea.getInitialSize() == new Size(7,11)
            eachArea.getTransformer() == transformer
            def commandDataList2 = eachArea.getCommandDataList()
            commandDataList2.size() == 1
            commandDataList2.get(0).startCellRef == new CellRef("Template!A9")
            commandDataList2.get(0).size == new Size(6,1)
            def command2 = commandDataList2.get(0).getCommand()
            command2.getName() == "each"
            command2 instanceof  EachCommand
            ((EachCommand)command2).items == "department.staff"
            ((EachCommand)command2).var == "employee"
            command2.getAreaList().size() == 1
            def eachArea2 = command2.getAreaList().get(0)
            eachArea2.getStartCellRef() == new CellRef("Template!A9")
            eachArea2.getInitialSize() == new Size(6,1)
            eachArea2.getTransformer() == transformer
            def commandDataList3 = eachArea2.getCommandDataList()
            commandDataList3.size() == 1
            commandDataList3.get(0).startCellRef == new CellRef("Template!A9")
            commandDataList3.get(0).getSize() == new Size(6,1)
            def command3 = commandDataList3.get(0).getCommand()
            command3.getName() == "if"
            command3 instanceof IfCommand
            ((IfCommand)command3).getCondition() == "employee.payment <= 2000"
            command3.getAreaList().size() == 2
            def ifArea = command3.getAreaList().get(0)
            def elseArea = command3.getAreaList().get(1)
            ifArea.getStartCellRef() == new CellRef("Template!A18")
            ifArea.getInitialSize() == new Size(6,1)
            ifArea.getCommandDataList().isEmpty()
            ifArea.getTransformer() == transformer
            elseArea.getStartCellRef() == new CellRef("Template!A9")
            elseArea.getInitialSize() == new Size(6,1)
            elseArea.getCommandDataList().isEmpty()
            elseArea.getTransformer() == transformer
            def xlsArea2 = xlsAreaList.get(1)
            xlsArea2.getStartCellRef() == new CellRef("Template!A2")
            xlsArea2.getInitialSize() == new Size(7,11)
            xlsArea2.getCommandDataList().isEmpty()
            xlsArea2.getTransformer() == transformer
    }
}
