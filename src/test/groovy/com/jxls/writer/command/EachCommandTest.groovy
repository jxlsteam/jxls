package com.jxls.writer.command

import spock.lang.Specification

import com.jxls.writer.Size
import com.jxls.writer.CellRef

/**
 * @author Leonid Vysochyn
 * Date: 1/18/12 6:00 PM
 */
class EachCommandTest extends Specification{
    def "test init"(){
        when:
            def area = new XlsArea(new CellRef(20, 20), new Size(4, 5))
            def eachCommand = new EachCommand( "x", "dataList", area)
        then:
             eachCommand.var == "x"
             eachCommand.items == "dataList"
             eachCommand.area == area
             eachCommand.direction == EachCommand.Direction.DOWN
    }

    def "test applyAt"(){
        given:
            def eachArea = Mock(Area)
            def eachCommand = new EachCommand( "x", "items", eachArea)
            def context = Mock(Context)
        when:
            eachCommand.applyAt(new CellRef("sheet2", 2, 2), context)
        then:
            context.toMap() >> ["items": [1,2,3,4]]
            1 * context.putVar("x", 1)
            1 * context.putVar("x", 2)
            1 * context.putVar("x", 3)
            1 * context.putVar("x", 4)
            4 * context.removeVar("x")
            1 * eachArea.applyAt(new CellRef("sheet2", 2, 2), context) >> new Size(3, 1)
            1 * eachArea.applyAt(new CellRef("sheet2", 3, 2), context) >> new Size(3, 2)
            1 * eachArea.applyAt(new CellRef("sheet2", 5, 2), context) >> new Size(4, 1)
            1 * eachArea.applyAt(new CellRef("sheet2", 6, 2), context) >> new Size(3, 1)
            0 * _._
    }
    
    def "test set direction"(){
        def area = Mock(Area)
        def eachCommand = new EachCommand("x", "dataList", area)
        when:
            eachCommand.setDirection(EachCommand.Direction.RIGHT)
        then:
            eachCommand.getDirection() == EachCommand.Direction.RIGHT
    }

    def "test applyAt with RIGHT direction"(){
        given:
            def eachArea = Mock(Area)
            def eachCommand = new EachCommand("x", "items", eachArea, EachCommand.Direction.RIGHT)
            def context = Mock(Context)
        when:
            eachArea.getInitialSize() >> new Size(3, 2)
            eachCommand.applyAt(new CellRef("sheet2", 1, 1), context)
        then:
            context.toMap() >> ["items": [1,2,3,4]]
            1 * context.putVar("x", 1)
            1 * context.putVar("x", 2)
            1 * context.putVar("x", 3)
            1 * context.putVar("x", 4)
            4 * context.removeVar("x")
            1 * eachArea.applyAt(new CellRef("sheet2", 1, 1), context) >> new Size(3, 1)
            1 * eachArea.applyAt(new CellRef("sheet2", 1, 4), context) >> new Size(3, 2)
            1 * eachArea.applyAt(new CellRef("sheet2", 1, 7), context) >> new Size(4, 1)
            1 * eachArea.applyAt(new CellRef("sheet2", 1, 11), context) >> new Size(3, 1)
            0 * _._
    }
}
