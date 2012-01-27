package com.jxls.writer.command

import spock.lang.Specification
import com.jxls.writer.Cell
import com.jxls.writer.Size

import com.jxls.writer.transform.Transformer
import com.jxls.writer.Pos

/**
 * @author Leonid Vysochyn
 * Date: 1/18/12 6:25 PM
 */
class BaseAreaTest extends Specification{
    def "test init"(){
        given:
            def transformer = Mock(Transformer)
        when:
            def area = new BaseArea(new Cell(1,1), new Size(5,5), transformer)
        then:
            assert area.startCell == new Cell(1,1)
            assert area.initialSize == new Size(5,5)
            assert area.transformer == transformer
    }
    
    def "test size"(){
        given:
            def area = new BaseArea(new Cell(1,1), new Size(10,15))
            def innerArea = Mock(Command)
            area.addCommand(new Pos(1,2), innerArea)
            def context = new Context()
        when:
            innerArea.getInitialSize() >> new Size(1,1)
            innerArea.getSize(context) >> new Size(3,5)
        then:
            assert area.getSize(context) == new Size(12, 19)
    }

    def "test applyAt with inner command"(){
        given:
            def area = new BaseArea(new Cell(1,1), new Size(10,15),Mock(Transformer))
            def innerArea = Mock(Command)
            def context = new Context()
            area.addCommand(new Pos(2,3), innerArea)
            innerArea.getInitialSize() >> new Size(2,3)
        when:
            area.applyAt(new Cell(4,5), context)
        then:
            1 * innerArea.applyAt(new Cell(6,8), context) >> new Size(2,5)
    }
    
    def "test applyAt for simple area"(){
        given:
            def area = new BaseArea(new Cell(1,1), new Size(2,3))
            def transformer = Mock(Transformer)
            def context = new Context()
            area.setTransformer(transformer)
        when:
            area.applyAt(new Cell(3,4), context)
        then:
            1 * transformer.transform(new Cell(1,1), new Cell(3,4), context)
            1 * transformer.transform(new Cell(1,2), new Cell(3,5), context)
            1 * transformer.transform(new Cell(1,3), new Cell(3,6), context)
            1 * transformer.transform(new Cell(2,1), new Cell(4,4), context)
            1 * transformer.transform(new Cell(2,2), new Cell(4,5), context)
            1 * transformer.transform(new Cell(2,3), new Cell(4,6), context)
            0 * _._
    }
    
    def "test applyAt with two inner commands"(){
        given:
            def transformer = Mock(Transformer)
            def area = new BaseArea(new Cell(1,1), new Size(10,15), transformer)
            def innerCommand1 = Mock(Command)
            def context = new Context()
            innerCommand1.getInitialSize() >> new Size(2,3)
            area.addCommand(new Pos(1,2), innerCommand1)
            def innerCommand2 = Mock(Command)
            innerCommand2.getInitialSize() >> new Size(4,5)
            area.addCommand(new Pos(0, 6), innerCommand2)
        when:
            area.applyAt(new Cell(4,5), context)
        then:
            1 * innerCommand1.applyAt(new Cell(5, 7), context) >> new Size(3,6)
            1 * innerCommand2.applyAt(new Cell(4, 14), context) >> new Size(4,3)
            1 * transformer.transform(new Cell(1,1), new Cell(4,5), context)
            1 * transformer.transform(new Cell(2,6), new Cell(5,13), context)
            1 * transformer.transform(new Cell(3,6), new Cell(6,13), context)
            1 * transformer.transform(new Cell(5,3), new Cell(9,7), context)
            1 * transformer.transform(new Cell(2,14), new Cell(5,19), context)
            1 * transformer.transform(new Cell(1,14), new Cell(4,19), context)
    }


}
