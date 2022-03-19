package org.jxls.builder.xml

import spock.lang.Specification
import org.jxls.common.CellRef
import org.jxls.common.Size
import org.jxls.command.EachCommand
import org.jxls.command.IfCommand
import org.jxls.transform.Transformer
import org.jxls.area.XlsArea

/**
 * @author Leonid Vysochyn
 * Date: 2/14/12 11:55 AM
 */
class XmlAreaBuilderTest extends Specification{
    def "test build"(){
        InputStream is = XmlAreaBuilderTest.class.getResourceAsStream("xlsarea.xml")
        Transformer transformer = Mock(Transformer)
        assert is != null
        when:
            def xlsAreaList = new XmlAreaBuilder(is, transformer).build()
        then:
            xlsAreaList.size() == 2
            def xlsArea = xlsAreaList.get(0)
            xlsArea != null
            xlsArea instanceof XlsArea
            xlsArea.getStartCellRef() == new CellRef("Template!A1")
            xlsArea.getSize() == new Size(7,15)
            xlsArea.getTransformer() == transformer
            def commandDataList1 = xlsArea.getCommandDataList()
            commandDataList1.size() == 1
            commandDataList1.get(0).startCellRef == new CellRef("Template!A2")
            commandDataList1.get(0).commandSize == new Size(6,11)
            def command1 = commandDataList1.get(0).getCommand()
            command1.getName() == "each"
            command1 instanceof EachCommand
            ((EachCommand)command1).items == "departments"
            ((EachCommand)command1).var == "department"
            command1.getAreaList().size() == 1
            def eachArea = command1.getAreaList().get(0)
            eachArea.getStartCellRef() == new CellRef("Template!A2")
            eachArea.getSize() == new Size(7,11)
            eachArea.getTransformer() == transformer
            def commandDataList2 = eachArea.getCommandDataList()
            commandDataList2.size() == 1
            commandDataList2.get(0).startCellRef == new CellRef("Template!A9")
            commandDataList2.get(0).commandSize == new Size(6,1)
            def command2 = commandDataList2.get(0).getCommand()
            command2.getName() == "each"
            command2 instanceof  EachCommand
            ((EachCommand)command2).items == "department.staff"
            ((EachCommand)command2).var == "employee"
            command2.getAreaList().size() == 1
            def eachArea2 = command2.getAreaList().get(0)
            eachArea2.getStartCellRef() == new CellRef("Template!A9")
            eachArea2.getSize() == new Size(6,1)
            eachArea2.getTransformer() == transformer
            def commandDataList3 = eachArea2.getCommandDataList()
            commandDataList3.size() == 1
            commandDataList3.get(0).startCellRef == new CellRef("Template!A9")
            commandDataList3.get(0).getCommandSize() == new Size(6,1)
            def command3 = commandDataList3.get(0).getCommand()
            command3.getName() == "if"
            command3 instanceof IfCommand
            ((IfCommand)command3).getCondition() == "employee.payment <= 2000"
            command3.getAreaList().size() == 2
            def ifArea = command3.getAreaList().get(0)
            def elseArea = command3.getAreaList().get(1)
            ifArea.getStartCellRef() == new CellRef("Template!A18")
            ifArea.getSize() == new Size(6,1)
            ifArea.getCommandDataList().isEmpty()
            ifArea.getTransformer() == transformer
            elseArea.getStartCellRef() == new CellRef("Template!A9")
            elseArea.getSize() == new Size(6,1)
            elseArea.getCommandDataList().isEmpty()
            elseArea.getTransformer() == transformer
            def xlsArea2 = xlsAreaList.get(1)
            xlsArea2.getStartCellRef() == new CellRef("Template!A2")
            xlsArea2.getSize() == new Size(7,11)
            xlsArea2.getCommandDataList().isEmpty()
            xlsArea2.getTransformer() == transformer
    }

    def "test build with custom action"(){
            InputStream is = XmlAreaBuilderTest.class.getResourceAsStream("useraction.xml")
            Transformer transformer = Mock(Transformer)
            assert is != null
        when:
            def xlsAreaList = new XmlAreaBuilder(is, transformer).build()
        then:
            xlsAreaList.size() == 1
            def xlsArea = xlsAreaList.get(0)
            xlsArea.getCommandDataList().size() == 1
            def eachCommand = xlsArea.getCommandDataList().get(0).getCommand();
            def eachArea = eachCommand.getAreaList().get(0);
            eachArea.getCommandDataList().size() == 1
            def customCommand = eachArea.getCommandDataList().get(0).getCommand()
            customCommand.getName() == "custom"
            customCommand instanceof CustomCommand
            ((CustomCommand)customCommand).getAttr() == "CustomValue"
            customCommand.getAreaList().size() == 1
            def customArea = customCommand.getAreaList().get(0)
            customArea.getCommandDataList().size() == 1
    }

    def "test build with user command"(){
            InputStream is = XmlAreaBuilderTest.class.getResourceAsStream("usercommand.xml")
            Transformer transformer = Mock(Transformer)
            assert is != null
        when:
            def xlsAreaList = new XmlAreaBuilder(is, transformer).build()
        then:
            xlsAreaList.size() == 1
            def xlsArea = xlsAreaList.get(0)
            xlsArea.getCommandDataList().size() == 1
            def eachCommand = xlsArea.getCommandDataList().get(0).getCommand();
            def eachArea = eachCommand.getAreaList().get(0);
            eachArea.getCommandDataList().size() == 1
            def customCommand = eachArea.getCommandDataList().get(0).getCommand()
            customCommand.getName() == "custom"
            customCommand instanceof CustomCommand
            ((CustomCommand)customCommand).getAttr() == "CustomValue"
            customCommand.getAreaList().size() == 1
            def customArea = customCommand.getAreaList().get(0)
            customArea.getCommandDataList().size() == 1
    }

    def "test build with grid command"(){
        InputStream is = XmlAreaBuilderTest.class.getResourceAsStream("grid.xml")
        Transformer transformer = Mock(Transformer)
        assert is != null
        when:
        def xlsAreaList = new XmlAreaBuilder(is, transformer).build()
        then:
        xlsAreaList.size() == 1
        def xlsArea = xlsAreaList.get(0)
        xlsArea.getCommandDataList().size() == 1
        def gridCommand = xlsArea.getCommandDataList().get(0).getCommand();
        def headerArea = gridCommand.getAreaList().get(0);
        headerArea.getCommandDataList().size() == 0
        def bodyArea = gridCommand.getAreaList().get(1)
        bodyArea.getCommandDataList().size() == 0
    }


}
