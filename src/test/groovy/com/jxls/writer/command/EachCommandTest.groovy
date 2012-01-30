package com.jxls.writer.command

import spock.lang.Specification
import com.jxls.writer.Cell
import com.jxls.writer.Size
import com.jxls.writer.Pos

/**
 * @author Leonid Vysochyn
 * Date: 1/18/12 6:00 PM
 */
class EachCommandTest extends Specification{
    def "test init"(){
        when:
            def area = new BaseArea(new Cell(20, 20), new Size(4, 5))
            def eachCommand = new EachCommand(new Cell(1,2), new Size(2,3), "x", "dataList", area)
        then:
            assert eachCommand.startCell == new Cell(1,2)
            assert eachCommand.initialSize == new Size(2,3)
            assert eachCommand.var == "x"
            assert eachCommand.items == "dataList"
            assert eachCommand.area == area
    }

    def "test size"(){
        given:
            def area = new BaseArea(new Cell(20, 20), new Size(4, 5))
            def eachCommand = new EachCommand(new Cell(1,2), new Size(2,3), "x", "dataList", area)
            def context = new Context()
        when:
            context.putVar("dataList", dataList)
        then:
            eachCommand.getSize(context) == size
        where:
            dataList        | size
            ['a', 'b', 'c'] | new Size(4, 15)
            ['x', 'y']      | new Size(4, 10)
            []              | new Size(0, 0)
    }
    
    def "test applyAt"(){
        given:
            def eachArea = Mock(Area)
            def eachCommand = new EachCommand(new Cell(1, 2), new Size(4, 2), "x", "items", eachArea)
            def context = Mock(Context)
        when:
            eachArea.getInitialSize() >> new Size(3,2)
            eachCommand.applyAt(new Cell(2,2,1), context)
        then:
            context.toMap() >> ["items": [1,2,3,4]]
            1 * context.putVar("x", 1)
            1 * context.putVar("x", 2)
            1 * context.putVar("x", 3)
            1 * context.putVar("x", 4)
            4 * context.removeVar("x")
            1 * eachArea.applyAt(new Cell(2,2,1), context) >> new Size(3, 1)
            1 * eachArea.applyAt(new Cell(2,3,1), context) >> new Size(3, 2)
            1 * eachArea.applyAt(new Cell(2,5,1), context) >> new Size(4, 1)
            1 * eachArea.applyAt(new Cell(2,6,1), context) >> new Size(3, 1)
            0 * _._
    }
}
