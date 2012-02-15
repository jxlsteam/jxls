package com.jxls.writer.command

import spock.lang.Specification

import com.jxls.writer.common.Size
import com.jxls.writer.common.CellRef
import com.jxls.writer.area.XlsArea
import com.jxls.writer.area.Area
import com.jxls.writer.common.Context

/**
 * @author Leonid Vysochyn
 * Date: 1/5/12
 */
class IfCommandTest extends Specification{
    def "test init" (){
        when:
            def ifArea = Mock(Area)
            def elseArea = Mock(Area)
            def ifCommand = new IfCommand("2*x + 5 > 10",
                    ifArea, elseArea );
        then:
             ifCommand.condition == "2*x + 5 > 10"
             ifCommand.ifArea == ifArea
             ifCommand.elseArea == elseArea
    }
    
    def "test add area"(){
        def ifArea = Mock(Area)
        def elseArea = Mock(Area)
        when:
            def ifCommand = new IfCommand("a > b")
            ifCommand.addArea(ifArea)
            ifCommand.addArea(elseArea)
        then:
            ifCommand.condition == "a > b"
            ifCommand.ifArea == ifArea
            ifCommand.elseArea == elseArea
    }

    def "test add excessive number of areas"(){
        def ifCommand = new IfCommand("a > b")
        ifCommand.addArea(Mock(Area))
        ifCommand.addArea(Mock(Area))
        when:
            ifCommand.addArea(Mock(Area))
        then:
            thrown(IllegalArgumentException)
    }

    def "test condition"(){
        given:
            def ifCommand = new IfCommand("2*x + 5 > 10",  new XlsArea(new CellRef(10, 5), new Size(5,5)), new XlsArea(new CellRef(10, 10), new Size(3,3)))
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
    
    def "test applyAt when condition is false"(){
        given:
            def ifArea = Mock(Area)
            def elseArea = Mock(Area)
            def ifCommand = new IfCommand("2*x + 5 > 10", ifArea, elseArea)
            def context = new Context()
        when:
            context.putVar("x", 2)
            ifCommand.applyAt(new CellRef(1, 1), context)
        then:
            1 * elseArea.applyAt(new CellRef(1, 1), context)
            0 * _._
    }

    def "test applyAt when condition is true"(){
        given:
            def ifArea = Mock(Area)
            def elseArea = Mock(Area)
            def ifCommand = new IfCommand("2*x + 5 > 10", ifArea, elseArea)
            def context = new Context()
        when:
            context.putVar("x", 3)
            ifCommand.applyAt(new CellRef(1, 1), context)
        then:
            1 * ifArea.applyAt(new CellRef(1, 1), context)
            0 * _._
    }

}
