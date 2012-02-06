package com.jxls.writer.command

import spock.lang.Specification

import com.jxls.writer.Size

import com.jxls.writer.transform.Transformer
import com.jxls.writer.Pos
import com.jxls.writer.CellData
import spock.lang.Ignore

/**
 * @author Leonid Vysochyn
 * Date: 1/18/12 6:25 PM
 */
class BaseAreaTest extends Specification{
    def "test init"(){
        given:
            def transformer = Mock(Transformer)
        when:
            def area = new BaseArea(new Pos(1, 1), new Size(5,5), transformer)
        then:
            area.startPos == new Pos(1, 1)
            area.initialSize == new Size(5,5)
            area.transformer == transformer
    }

    def "test applyAt with inner command"(){
        given:
            def area = new BaseArea(new Pos(1, 1), new Size(10,15),Mock(Transformer))
            def innerCommand = Mock(Command)
            def context = new Context()
            area.addCommand(new Pos(3, 2), innerCommand)
            innerCommand.getInitialSize() >> new Size(2,3)
        when:
            area.applyAt(new Pos(1, 5, 4), context)
        then:
            1 * innerCommand.applyAt(new Pos(1, 8, 6), context) >> new Size(2,5)
    }
    
    def "test applyAt for simple area"(){
        given:
            def area = new BaseArea(new Pos(1, 1), new Size(2,3))
            def transformer = Mock(Transformer)
            def context = new Context()
            area.setTransformer(transformer)
        when:
            area.applyAt(new Pos(4, 3), context)
        then:
            1 * transformer.transform(new Pos(1, 1), new Pos(4, 3), context)
            1 * transformer.transform(new Pos(2, 1), new Pos(5, 3), context)
            1 * transformer.transform(new Pos(3, 1), new Pos(6, 3), context)
            1 * transformer.transform(new Pos(1, 2), new Pos(4, 4), context)
            1 * transformer.transform(new Pos(2, 2), new Pos(5, 4), context)
            1 * transformer.transform(new Pos(3, 2), new Pos(6, 4), context)
            0 * _._
    }
    
    def "test applyAt for another sheet"(){
        given:
            def area = new BaseArea(new Pos(1, 1), new Size(2,3))
            def transformer = Mock(Transformer)
            def context = new Context()
            area.setTransformer(transformer)
        when:
            area.applyAt(new Pos(1, 4, 3), context)
        then:
            1 * transformer.transform(new Pos(1, 1), new Pos(1, 4, 3), context)
            1 * transformer.transform(new Pos(2, 1), new Pos(1, 5, 3), context)
            1 * transformer.transform(new Pos(3, 1), new Pos(1, 6, 3), context)
            1 * transformer.transform(new Pos(1, 2), new Pos(1, 4, 4), context)
            1 * transformer.transform(new Pos(2, 2), new Pos(1, 5, 4), context)
            1 * transformer.transform(new Pos(3, 2), new Pos(1, 6, 4), context)
            0 * _._
    }

    def "test applyAt with two inner commands"(){
        given:
            def transformer = Mock(Transformer)
            def area = new BaseArea(new Pos(1, 1), new Size(10,15), transformer)
            def innerCommand1 = Mock(Command)
            def context = new Context()
            innerCommand1.getInitialSize() >> new Size(2,3)
            area.addCommand(new Pos(2, 1), innerCommand1)
            def innerCommand2 = Mock(Command)
            innerCommand2.getInitialSize() >> new Size(4,5)
            area.addCommand(new Pos(6, 0), innerCommand2)
        when:
            area.applyAt(new Pos(5, 4), context)
        then:
            1 * innerCommand1.applyAt(new Pos(7, 5), context) >> new Size(3,6)
            1 * innerCommand2.applyAt(new Pos(14, 4), context) >> new Size(4,3)
            1 * transformer.transform(new Pos(1, 1), new Pos(5, 4), context)
            1 * transformer.transform(new Pos(6, 2), new Pos(13, 5), context)
            1 * transformer.transform(new Pos(6, 3), new Pos(13, 6), context)
            1 * transformer.transform(new Pos(3, 5), new Pos(7, 9), context)
            1 * transformer.transform(new Pos(14, 2), new Pos(19, 5), context)
            1 * transformer.transform(new Pos(14, 1), new Pos(19, 4), context)
    }

    def "test applyAt multiple times"(){
        given:
            def transformer = Mock(Transformer)
            def area = new BaseArea(new Pos(1, 1), new Size(2,1), transformer)
            Context context1 = new Context()
            context1.putVar("x", 1)
            Context context2 = new Context()
            context2.putVar("x", 2)
        when:
            area.applyAt(new Pos(2, 2), context1)
            area.applyAt(new Pos(10, 2), context2)
        then:
            1 * transformer.transform(new Pos(1, 1), new Pos(2, 2), context1)
            1 * transformer.transform(new Pos(1, 2), new Pos(2, 3), context1)
            1 * transformer.transform(new Pos(1, 1), new Pos(10, 2), context2)
            1 * transformer.transform(new Pos(1, 2), new Pos(10, 3), context2)
            0 * _._
    }

    def "test formulas transformation"(){
        given:
            def transformer = Mock(Transformer)
            def area = new BaseArea(new Pos(1, 1), new Size(3,3), transformer)
            Context context = new Context()
            context.putVar("x", 1)
        when:
            area.processFormulas()
        then:
            1 * transformer.getFormulaCells() >> [new CellData(0, 1, 1, CellData.CellType.FORMULA, "A1+B3"),
                    new CellData(0, 3, 1, CellData.CellType.FORMULA, "D20 * E30"),
                    new CellData(0, 2, 2, CellData.CellType.STRING, '$[SUM(F7)]')]
            1 * transformer.getTargetPos(new Pos(0,0,0)) >> [new Pos(0,0,0)]
            1 * transformer.getTargetPos(new Pos(0,2,1)) >> [new Pos("C5")]
            1 * transformer.getTargetPos(new Pos(0,1,1)) >> [new Pos(1,11,12)]

            1 * transformer.getTargetPos(new Pos(0,19,3)) >> [new Pos("K10")]
            1 * transformer.getTargetPos(new Pos(0,29,4)) >> [new Pos("I20")]
            1 * transformer.getTargetPos(new Pos(0,6,5)) >> ([new Pos("R77"), new Pos("R78"), new Pos("R79")])
            1 * transformer.getTargetPos(new Pos(0,2,2)) >> [new Pos(0,31,35)]
            1 * transformer.getTargetPos(new Pos(0,3,1)) >> [new Pos(0,22,23)]
            1 * transformer.updateFormula(new Pos(0,22,23), "K10 * I20")
            1 * transformer.updateFormula(new Pos(1, 11, 12), "A1+C5")
            1 * transformer.updateFormula(new Pos(31, 35), "SUM(R77,R78,R79)")
            0 * _._
    }

    def "test formula processing when transforming into multiple cells"(){
        given:
            def transformer = Mock(Transformer)
            def area = new BaseArea(new Pos(1, 1), new Size(3,3), transformer)
            Context context = new Context()
            context.putVar("x", 1)
        when:
            area.processFormulas()
        then:
            1 * transformer.getFormulaCells() >> [new CellData(0, 1, 2, CellData.CellType.FORMULA, "A1+B2+C5")]
            1 * transformer.getTargetPos(new Pos(0,0,0)) >> [new Pos("A1"), new Pos("A3"), new Pos("A5")]
            1 * transformer.getTargetPos(new Pos(0,1,1)) >> [new Pos("B2"), new Pos("B4"), new Pos("B6")]
            1 * transformer.getTargetPos(new Pos(0,4,2)) >> [new Pos("C10")]
            1 * transformer.getTargetPos(new Pos(0,1,2)) >> [new Pos(0,1,2), new Pos(0, 3, 2), new Pos(0,5,2)]
            1 * transformer.updateFormula(new Pos(0,1,2), "A1+B2+C10")
            1 * transformer.updateFormula(new Pos(0,3,2), "A3+B4+C10")
            1 * transformer.updateFormula(new Pos(0,5,2), "A5+B6+C10")
            0 * _._
    }
    @Ignore
    def "test formula processing for nested formulas"(){
        given:
            def transformer = Mock(Transformer)
            def area = new BaseArea(new Pos(1, 1), new Size(5,5), transformer)
            Context context = new Context()
            context.putVar("x", 1)
        when:
            area.processFormulas()
        then:
            1 * transformer.getFormulaCells() >> [new CellData(0, 1, 2, CellData.CellType.FORMULA, "SUM(B1,B20)")]
            1 * transformer.getTargetPos(new Pos(0,1,2)) >> [new Pos(0,5,2), new Pos(0, 10, 2), new Pos(0, 15, 2)]
            1 * transformer.getTargetPos(new Pos(0,0,1)) >> [new Pos("B2"), new Pos("B4"),  new Pos("B9"), new Pos("B10"), new Pos("B15")]
            1 * transformer.getTargetPos(new Pos(0,19,1)) >> [new Pos("B1"), new Pos("B3"), new Pos("B8"), new Pos("B16")]
            1 * transformer.updateFormula(new Pos(0,5,2), "SUM(B1,B2,B3,B4)")
            1 * transformer.updateFormula(new Pos(0,10,2), "SUM(B8,B9,B10)")
            1 * transformer.updateFormula(new Pos(0,15,2), "SUM(B15,B16)")
    }


}
