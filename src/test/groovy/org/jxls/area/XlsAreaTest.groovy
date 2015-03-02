package org.jxls.area

import spock.lang.Specification

import org.jxls.common.Size

import org.jxls.transform.Transformer

import org.jxls.common.CellData
import org.jxls.common.CellRef
import org.jxls.common.AreaRef
import org.jxls.common.Context
import org.jxls.command.Command
import org.jxls.common.AreaListener

/**
 * @author Leonid Vysochyn
 * Date: 1/18/12 6:25 PM
 */
class XlsAreaTest extends Specification{
    def "test create"(){
        given:
            def transformer = Mock(Transformer)
        when:
            def area = new XlsArea(new CellRef(1, 1), new Size(5,5), transformer)
        then:
            area.startCellRef == new CellRef(1, 1)
            area.size == new Size(5,5)
            area.transformer == transformer
            !area.clearCellsBeforeApply
    }

    def "test create from two cell refs"(){
        def transformer = Mock(Transformer)
        when:
            def area = new XlsArea(new CellRef("sheet1!A1"), new CellRef("sheet1!D3"), transformer)
        then:
            area.startCellRef == new CellRef("sheet1!A1")
            area.size == new Size(4,3)
            area.transformer == transformer
    }
    
    def "test create from area ref"(){
        when:
            def area = new XlsArea("sheet1!A1:D3", Mock(Transformer))
        then:
            area.startCellRef == new CellRef("sheet1!A1")
            area.size == new Size(4,3)
    }
    
    def "test add command with cell ref, size"(){
        given:
            def xlsArea = new XlsArea("sheet1!A1:C4", Mock(Transformer))
            def command = Mock(Command)
        when:
            xlsArea.addCommand(new AreaRef( new CellRef("sheet1!A2"), new Size(2,1)), command)
        then:
            xlsArea.getCommandDataList().size() == 1
            CommandData commandData = xlsArea.getCommandDataList().get(0)
            commandData.getStartCellRef() == new CellRef("sheet1!A2")
            commandData.getSize() == new Size(2,1)
            commandData.getCommand() == command
    }

    def "test add command with area ref instance"(){
        given:
        def xlsArea = new XlsArea("sheet1!A1:C4", Mock(Transformer))
        def command = Mock(Command)
        when:
        xlsArea.addCommand(new AreaRef("sheet1!A2:B3"), command)
        then:
        xlsArea.getCommandDataList().size() == 1
        CommandData commandData = xlsArea.getCommandDataList().get(0)
        commandData.getStartCellRef() == new CellRef("sheet1!A2")
        commandData.getSize() == new Size(2,2)
        commandData.getCommand() == command
    }

    def "test add command with area ref"(){
        given:
        def xlsArea = new XlsArea("sheet1!A1:C4", Mock(Transformer))
        def command = Mock(Command)
        when:
        xlsArea.addCommand("sheet1!A2:B3", command)
        then:
        xlsArea.getCommandDataList().size() == 1
        CommandData commandData = xlsArea.getCommandDataList().get(0)
        commandData.getStartCellRef() == new CellRef("sheet1!A2")
        commandData.getSize() == new Size(2,2)
        commandData.getCommand() == command
    }
    
    def "test clear cells"(){
        given:
            def transformer = Mock(Transformer)
            def area = new XlsArea(new CellRef("sheet1", 1, 1), new Size(2,2), transformer)
        when:
            area.clearCells()
        then:
            1 * transformer.clearCell(new CellRef("sheet1", 1, 1))
            1 * transformer.clearCell(new CellRef("sheet1", 1, 2))
            1 * transformer.clearCell(new CellRef("sheet1", 2, 1))
            1 * transformer.clearCell(new CellRef("sheet1", 2, 2))
    }

    def "test applyAt with default clearCells flag value set to true"(){
        given:
            def transformer = Mock(Transformer)
            def area = new XlsArea(new CellRef("sheet1", 1, 1), new Size(2,2), transformer)
            def context = new Context()
        when:
            area.setClearCellsBeforeApply(true)
            area.applyAt(new CellRef("sheet1", 3,3), context)
        then:
            1 * transformer.transform(new CellRef("sheet1", 1, 1), new CellRef("sheet1", 3, 3), context)
            1 * transformer.transform(new CellRef("sheet1", 1, 2), new CellRef("sheet1", 3, 4), context)
            1 * transformer.transform(new CellRef("sheet1", 2, 1), new CellRef("sheet1", 4, 3), context)
            1 * transformer.transform(new CellRef("sheet1", 2, 2), new CellRef("sheet1", 4, 4), context)

            1 * transformer.clearCell(new CellRef("sheet1", 1, 1))
            1 * transformer.clearCell(new CellRef("sheet1", 1, 2))
            1 * transformer.clearCell(new CellRef("sheet1", 2, 1))
            1 * transformer.clearCell(new CellRef("sheet1", 2, 2))
            0 * _._
    }

    def "test applyAt with clearCells flag value set to false"(){
        given:
            def transformer = Mock(Transformer)
            def area = new XlsArea(new CellRef("sheet1", 1, 1), new Size(2,2), transformer)
            def context = new Context()
        when:
            area.setClearCellsBeforeApply(false)
            area.applyAt(new CellRef("sheet1", 3,3), context)
        then:
            1 * transformer.transform(new CellRef("sheet1", 1, 1), new CellRef("sheet1", 3, 3), context)
            1 * transformer.transform(new CellRef("sheet1", 1, 2), new CellRef("sheet1", 3, 4), context)
            1 * transformer.transform(new CellRef("sheet1", 2, 1), new CellRef("sheet1", 4, 3), context)
            1 * transformer.transform(new CellRef("sheet1", 2, 2), new CellRef("sheet1", 4, 4), context)
            0 * _._
    }

    def "test applyAt with inner command"(){
        given:
            def area = new XlsArea(new CellRef("sheet1", 1, 1), new Size(10,15),Mock(Transformer))
            def innerCommand = Mock(Command)
            def context = new Context()
            area.addCommand(new AreaRef(new CellRef("sheet1",3, 2), new Size(2,3)), innerCommand)
        when:
            area.applyAt(new CellRef("sheet2", 5, 4), context)
        then:
            1 * innerCommand.applyAt(new CellRef("sheet2", 7, 5), context) >> new Size(2,5)
    }
    
    def "test applyAt for simple area"(){
        given:
            def area = new XlsArea(new CellRef("sheet1", 1, 1), new Size(2,3))
            def transformer = Mock(Transformer)
            def context = new Context()
            area.setTransformer(transformer)
        when:
            area.applyAt(new CellRef("sheet1", 4, 3), context)
        then:
            1 * transformer.transform(new CellRef("sheet1",1, 1), new CellRef("sheet1",4, 3), context)
            1 * transformer.transform(new CellRef("sheet1",2, 1), new CellRef("sheet1",5, 3), context)
            1 * transformer.transform(new CellRef("sheet1",3, 1), new CellRef("sheet1",6, 3), context)
            1 * transformer.transform(new CellRef("sheet1",1, 2), new CellRef("sheet1",4, 4), context)
            1 * transformer.transform(new CellRef("sheet1",2, 2), new CellRef("sheet1",5, 4), context)
            1 * transformer.transform(new CellRef("sheet1",3, 2), new CellRef("sheet1",6, 4), context)
    }
    
    def "test applyAt for another sheet"(){
        given:
            def area = new XlsArea(new CellRef("sheet1",1, 1), new Size(2,3))
            def transformer = Mock(Transformer)
            def context = new Context()
            area.setTransformer(transformer)
        when:
            area.applyAt(new CellRef("sheet2", 4, 3), context)
        then:
            1 * transformer.transform(new CellRef("sheet1",1, 1), new CellRef("sheet2", 4, 3), context)
            1 * transformer.transform(new CellRef("sheet1",2, 1), new CellRef("sheet2", 5, 3), context)
            1 * transformer.transform(new CellRef("sheet1",3, 1), new CellRef("sheet2", 6, 3), context)
            1 * transformer.transform(new CellRef("sheet1",1, 2), new CellRef("sheet2", 4, 4), context)
            1 * transformer.transform(new CellRef("sheet1",2, 2), new CellRef("sheet2", 5, 4), context)
            1 * transformer.transform(new CellRef("sheet1",3, 2), new CellRef("sheet2", 6, 4), context)
    }

    def "test applyAt with two inner commands"(){
        given:
        def transformer = Mock(Transformer)
        def area = new XlsArea(new CellRef("sheet1",1, 1), new Size(10,15), transformer)
        def innerCommand1 = Mock(Command)
        def context = new Context()
        area.addCommand(new AreaRef(new CellRef("sheet1",3, 2), new Size(2,3)), innerCommand1)
        def innerCommand2 = Mock(Command)
        area.addCommand(new AreaRef(new CellRef("sheet1",7, 1), new Size(4,5)), innerCommand2)
        when:
        area.applyAt(new CellRef("sheet1",5, 4), context)
        then:
        1 * innerCommand1.applyAt(new CellRef("sheet1",7, 5), context) >> new Size(3,6)
        1 * innerCommand2.applyAt(new CellRef("sheet1",14, 4), context) >> new Size(4,3)
        1 * transformer.transform(new CellRef("sheet1",1, 1), new CellRef("sheet1",5, 4), context)
        1 * transformer.transform(new CellRef("sheet1",6, 2), new CellRef("sheet1",13, 5), context)
        1 * transformer.transform(new CellRef("sheet1",6, 3), new CellRef("sheet1",13, 6), context)
        1 * transformer.transform(new CellRef("sheet1",3, 5), new CellRef("sheet1",7, 9), context)
        1 * transformer.transform(new CellRef("sheet1",14, 2), new CellRef("sheet1",19, 5), context)
        1 * transformer.transform(new CellRef("sheet1",14, 1), new CellRef("sheet1",19, 4), context)
    }

    def "test applyAt multiple times"(){
        given:
            def transformer = Mock(Transformer)
            def area = new XlsArea(new CellRef("sheet1",1, 1), new Size(2,1), transformer)
            Context context1 = new Context()
            context1.putVar("x", 1)
            Context context2 = new Context()
            context2.putVar("x", 2)
        when:
            area.applyAt(new CellRef("sheet1",2, 2), context1)
            area.applyAt(new CellRef("sheet1",10, 2), context2)
        then:
            1 * transformer.transform(new CellRef("sheet1",1, 1), new CellRef("sheet1",2, 2), context1)
            1 * transformer.transform(new CellRef("sheet1",1, 2), new CellRef("sheet1",2, 3), context1)
            1 * transformer.transform(new CellRef("sheet1",1, 1), new CellRef("sheet1",10, 2), context2)
            1 * transformer.transform(new CellRef("sheet1",1, 2), new CellRef("sheet1",10, 3), context2)
    }

    def "test formulas transformation"(){
        given:
            def transformer = Mock(Transformer)
            def area = new XlsArea(new CellRef("sheet1",1, 1), new Size(3,3), transformer)
            Context context = new Context()
            context.putVar("x", 1)
        when:
            area.processFormulas()
        then:
            1 * transformer.getFormulaCells() >> [new CellData("sheet1", 1, 1, CellData.CellType.FORMULA, "A1+B3"),
                    new CellData("sheet1", 3, 1, CellData.CellType.FORMULA, "D20 * E30"),
                    new CellData("sheet1", 2, 2, CellData.CellType.STRING, '$[SUM(F7)]')]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",0,0)) >> [new CellRef("sheet1",0,0)]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",2,1)) >> [new CellRef("sheet1!C5")]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",1,1)) >> [new CellRef("sheet2",11,12)]

            1 * transformer.getTargetCellRef(new CellRef("sheet1",19,3)) >> [new CellRef("sheet1!K10")]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",29,4)) >> [new CellRef("sheet1!I20")]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",6,5)) >> ([new CellRef("sheet1!R77"), new CellRef("sheet1!R78"), new CellRef("sheet1!R79")])
            1 * transformer.getTargetCellRef(new CellRef("sheet1",2,2)) >> [new CellRef("sheet1",31,35)]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",3,1)) >> [new CellRef("sheet1",22,23)]
            1 * transformer.setFormula(new CellRef("sheet1",22,23), "K10 * I20")
            1 * transformer.setFormula(new CellRef("sheet2", 11, 12), "sheet1!A1+sheet1!C5")
            1 * transformer.setFormula(new CellRef("sheet1", 31, 35), "SUM(R77:R79)")
            0 * _._
    }

    def "test formula processing when transforming into multiple cells"(){
        given:
            def transformer = Mock(Transformer)
            def area = new XlsArea(new CellRef("sheet1",1, 1), new Size(3,3), transformer)
            Context context = new Context()
            context.putVar("x", 1)
        when:
            area.processFormulas()
        then:
            1 * transformer.getFormulaCells() >> [new CellData("sheet1", 1, 2, CellData.CellType.FORMULA, "A1+B2+C5")]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",0,0)) >> [new CellRef("sheet1!A1"), new CellRef("sheet1!A3"), new CellRef("sheet1!A5")]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",1,1)) >> [new CellRef("sheet1!B2"), new CellRef("sheet1!B4"), new CellRef("sheet1!B6")]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",4,2)) >> [new CellRef("sheet1!C10")]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",1,2)) >> [new CellRef("sheet1",1,2), new CellRef("sheet1", 3, 2), new CellRef("sheet1",5,2)]
            1 * transformer.setFormula(new CellRef("sheet1",1,2), "A1+B2+C10")
            1 * transformer.setFormula(new CellRef("sheet1",3,2), "A3+B4+C10")
            1 * transformer.setFormula(new CellRef("sheet1",5,2), "A5+B6+C10")
            0 * _._
    }

    def "test formula processing for nested cells formulas"(){
        given:
            def transformer = Mock(Transformer)
            def area = new XlsArea(new CellRef("sheet1",1, 1), new Size(5,5), transformer)
            Context context = new Context()
            context.putVar("x", 1)
        when:
            area.processFormulas()
        then:
            1 * transformer.getFormulaCells() >> [new CellData("sheet1", 1, 2, CellData.CellType.FORMULA, "SUM(B1)")]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",1,2)) >> [new CellRef("sheet1",5,2), new CellRef("sheet1", 10, 2), new CellRef("sheet1", 15, 2)]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",0,1)) >> [new CellRef("sheet1!B2"), new CellRef("sheet1!B4"),  new CellRef("sheet1!B9"), new CellRef("sheet1!B10"), new CellRef("sheet1!B15"), new CellRef("sheet1!B1"),
                    new CellRef("sheet1!B3"), new CellRef("sheet1!B8"), new CellRef("sheet1!B16")]
            1 * transformer.setFormula(new CellRef("sheet1",5,2), "SUM(B1:B4)")
            1 * transformer.setFormula(new CellRef("sheet1",10,2), "SUM(B8:B10)")
            1 * transformer.setFormula(new CellRef("sheet1",15,2), "SUM(B15:B16)")
    }

    def "test formula processing for column formulas with joint cells"(){
        given:
            def transformer = Mock(Transformer)
            def area = new XlsArea(new CellRef("sheet1",1, 1), new Size(5,5), transformer)
            Context context = new Context()
            context.putVar("x", 1)
        when:
            area.processFormulas()
        then:
            1 * transformer.getFormulaCells() >> [new CellData("sheet1", 1, 2, CellData.CellType.FORMULA, "SUM(U_(B1,B20))")]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",1,2)) >> [new CellRef("sheet1",5,2), new CellRef("sheet1", 10, 2), new CellRef("sheet1", 15, 2)]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",0,1)) >> [new CellRef("sheet1!B2"), new CellRef("sheet1!B4"),  new CellRef("sheet1!B9"), new CellRef("sheet1!B10"), new CellRef("sheet1!B15")]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",19,1)) >> [new CellRef("sheet1!B1"), new CellRef("sheet1!B3"), new CellRef("sheet1!B8"), new CellRef("sheet1!B16")]
            1 * transformer.setFormula(new CellRef("sheet1",5,2), "SUM(B1:B4)")
            1 * transformer.setFormula(new CellRef("sheet1",10,2), "SUM(B8:B10)")
            1 * transformer.setFormula(new CellRef("sheet1",15,2), "SUM(B15:B16)")
    }

    def "test formula processing for row formulas with joint cells"(){
        given:
            def transformer = Mock(Transformer)
            def area = new XlsArea(new CellRef("sheet1",1, 1), new Size(5,5), transformer)
            Context context = new Context()
            context.putVar("x", 1)
        when:
            area.processFormulas()
        then:
            1 * transformer.getFormulaCells() >> [new CellData("sheet1", 1, 2, CellData.CellType.FORMULA, "SUM(U_(B1,B20))")]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",1,2)) >> [new CellRef("sheet1",5,2), new CellRef("sheet1", 10, 2), new CellRef("sheet1", 15, 2)]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",0,1)) >> [new CellRef("sheet1!B2"), new CellRef("sheet1!A4"),  new CellRef("sheet1!D2"), new CellRef("sheet1!C2"), new CellRef("sheet1!B3")]
            1 * transformer.getTargetCellRef(new CellRef("sheet1",19,1)) >> [new CellRef("sheet1!B4"), new CellRef("sheet1!E2")]
            1 * transformer.setFormula(new CellRef("sheet1",5,2), "SUM(B2:E2)")
            1 * transformer.setFormula(new CellRef("sheet1",10,2), "SUM(B3)")
            1 * transformer.setFormula(new CellRef("sheet1",15,2), "SUM(A4:B4)")
    }

    def "test formulas with other sheet references"(){
        given:
        def transformer = Mock(Transformer)
        def area = new XlsArea(new CellRef("sheet1",1, 1), new Size(5,5), transformer)
        Context context = new Context()
        context.putVar("x", 1)
        when:
        area.processFormulas()
        then:
        1 * transformer.getFormulaCells() >> [new CellData("sheet1", 2, 1, CellData.CellType.FORMULA, '''A1+Sheet2!A1 + 'Sheet 3'!B1 + Sheet2!B1 * '$ & test@.'!A5''')]
        1 * transformer.getTargetCellRef(new CellRef("sheet1",2,1)) >> [new CellRef("sheet1",5,5)]
        1 * transformer.getTargetCellRef(new CellRef("sheet1",0,0)) >> [new CellRef("sheet1",9,5)]
        1 * transformer.getTargetCellRef(new CellRef("Sheet2",0,0)) >> []
        1 * transformer.getTargetCellRef(new CellRef("Sheet 3",0,1)) >> [new CellRef("sheet1",5,2)]
        1 * transformer.getTargetCellRef(new CellRef("Sheet2",0,1)) >> [new CellRef("Sheet 3",0,0)]
        1 * transformer.getTargetCellRef(new CellRef('$ & test@.', 4, 0)) >> [ new CellRef('$ & test@.',4,1)]
        1 * transformer.setFormula(new CellRef("sheet1",5,5), '''F10+Sheet2!A1 + C6 + 'Sheet 3'!A1 * '$ & test@.'!B5''')
        0 * _._
    }
    
    def "test add/get area listener"(){
        given:
            def area = new XlsArea("sheet1!A1:C3", Mock(Transformer))
            def listener1 = Mock(AreaListener)
            def listener2 = Mock(AreaListener)
        when:
            area.addAreaListener(listener1)
            area.addAreaListener(listener2)
        then:
            area.getAreaListeners() == [listener1, listener2]
    }
    
    def "test invoke area listener"(){
        given:
            def transformer = Mock(Transformer)
            def area = new XlsArea(new CellRef("sheet1",1,1), new Size(2,2), transformer)
            def listener1 = Mock(AreaListener)
            def context1 = new Context()
            def context2 = new Context()
        when:
            area.addAreaListener(listener1)
            area.applyAt(new CellRef("sheet1", 3,3), context1)
            area.applyAt(new CellRef("sheet2", 1,2), context2)
        then:
            1 * listener1.beforeApplyAtCell(new CellRef("sheet1", 3,3), context1)
            1 * listener1.afterApplyAtCell(new CellRef("sheet1", 3,3), context1)
            1 * listener1.beforeApplyAtCell(new CellRef("sheet2", 1,2), context2)
            1 * listener1.afterApplyAtCell(new CellRef("sheet2", 1,2), context2)
            1 * listener1.beforeTransformCell(new CellRef("sheet1", 1,1), new CellRef("sheet1", 3,3), context1)
            1 * listener1.beforeTransformCell(new CellRef("sheet1", 1,2), new CellRef("sheet1", 3,4), context1)
            1 * listener1.beforeTransformCell(new CellRef("sheet1", 2,1), new CellRef("sheet1", 4,3), context1)
            1 * listener1.beforeTransformCell(new CellRef("sheet1", 2,2), new CellRef("sheet1", 4,4), context1)
            1 * listener1.afterTransformCell(new CellRef("sheet1", 1,1), new CellRef("sheet1", 3,3), context1)
            1 * listener1.afterTransformCell(new CellRef("sheet1", 1,2), new CellRef("sheet1", 3,4), context1)
            1 * listener1.afterTransformCell(new CellRef("sheet1", 2,1), new CellRef("sheet1", 4,3), context1)
            1 * listener1.afterTransformCell(new CellRef("sheet1", 2,2), new CellRef("sheet1", 4,4), context1)
            1 * listener1.beforeTransformCell(new CellRef("sheet1", 1,1), new CellRef("sheet2", 1,2), context2)
            1 * listener1.beforeTransformCell(new CellRef("sheet1", 1,2), new CellRef("sheet2", 1,3), context2)
            1 * listener1.beforeTransformCell(new CellRef("sheet1", 2,1), new CellRef("sheet2", 2,2), context2)
            1 * listener1.beforeTransformCell(new CellRef("sheet1", 2,2), new CellRef("sheet2", 2,3), context2)
            1 * listener1.afterTransformCell(new CellRef("sheet1", 1,1), new CellRef("sheet2", 1,2), context2)
            1 * listener1.afterTransformCell(new CellRef("sheet1", 1,2), new CellRef("sheet2", 1,3), context2)
            1 * listener1.afterTransformCell(new CellRef("sheet1", 2,1), new CellRef("sheet2", 2,2), context2)
            1 * listener1.afterTransformCell(new CellRef("sheet1", 2,2), new CellRef("sheet2", 2,3), context2)
            0 * listener1._
    }
    
    def "test adding commands outside area"(){
        def transformer = Mock(Transformer)
        def area = new XlsArea(new AreaRef("sheet1!C5:F10"), transformer)
        def command1 = Mock(Command)
        command1.getName() >> "command 1"
        when:
            area.addCommand(new AreaRef("sheet1!A4:C4"), command1)
        then:
            def e = thrown(IllegalArgumentException)
            e.cause == null
            e.message == "Cannot add command 'command 1' to area sheet1!C5:F10 at sheet1!A4:C4"
    }

}
