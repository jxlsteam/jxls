package com.jxls.writer.command

import spock.lang.Specification

import com.jxls.writer.Pos
import com.jxls.writer.Size

/**
 * @author Leonid Vysochyn
 * Date: 1/5/12
 */
class IfCommandTest extends Specification{
    def "test init" (){
        when:
            def ifArea = new BaseCommand(new Pos(5, 10), new Size(5,5))
            def elseArea = new BaseCommand(new Pos(10,10), new Size(3,3))
            def ifCommand = new IfCommand("2*x + 5 > 10", new Pos(2, 4), new Size(1,1),
                    ifArea, elseArea );
        then:
            assert ifCommand.initialSize == new Size(1,1)
            assert ifCommand.pos == new Pos(2,4)
            assert ifCommand.condition == "2*x + 5 > 10"
            assert ifCommand.ifArea == ifArea
            assert ifCommand.elseArea == elseArea
    }

    def "test condition"(){
        given:
            def ifCommand = new IfCommand("2*x + 5 > 10", new Pos(2,4), new Size(1,1), new BaseCommand(new Pos(5, 10), new Size(5,5)), new BaseCommand(new Pos(10,10), new Size(3,3)))
            def context = new Context()
        when:
            context.putVar("x", xValue)
        then:
            ifCommand.isConditionTrue(context) == result
        where:
            xValue  | result
            2       | false
            3       | true
    }
    
    def "test size" (){
        given:
            def ifCommand = new IfCommand("2*x + 5 > 10", new Pos(2,4), new Size(1,1), new BaseCommand(new Pos(5, 10), new Size(5,5)), new BaseCommand(new Pos(10,10), new Size(3,3)))
            def context = new Context()
        when:
            context.putVar("x", xValue)
        then:
            ifCommand.getSize(context) == new Size(width, height)
        where:
            xValue | width | height
            2      | 3     | 3
            3      | 5     | 5
    }
    
    def "test size without else section"(){
        given:
            def ifCommand = new IfCommand("2*x + 5 > 10", new Pos(2,4), new Size(1,1), new BaseCommand(new Pos(5, 10), new Size(5,5)))
            def context = new Context()
        when:
            context.putVar("x", xValue)
        then:
            ifCommand.getSize(context) == new Size(width, height)
        where:
            xValue | width | height
            2      | 0     | 0
            3      | 5     | 5
    }

    def "test applyAt when condition is false"(){
        given:
            def ifArea = Mock(Command)
            def elseArea = Mock(Command)
            def ifCommand = new IfCommand("2*x + 5 > 10", new Pos(2,4), new Size(1,1), ifArea, elseArea)
            def context = new Context()
        when:
            context.putVar("x", 2)
            ifCommand.applyAt(new Pos(1,1), context)
        then:
            1 * elseArea.applyAt(new Pos(1,1), context)
            0 * _._
    }

    def "test applyAt when condition is true"(){
        given:
            def ifArea = Mock(Command)
            def elseArea = Mock(Command)
            def ifCommand = new IfCommand("2*x + 5 > 10", new Pos(2,4), new Size(1,1), ifArea, elseArea)
            def context = new Context()
        when:
            context.putVar("x", 3)
            ifCommand.applyAt(new Pos(1,1), context)
        then:
            1 * ifArea.applyAt(new Pos(1,1), context)
            0 * _._
    }

}
