package com.jxls.writer.command

import spock.lang.Specification
import com.jxls.writer.Pos
import com.jxls.writer.Size

/**
 * @author Leonid Vysochyn
 * Date: 1/18/12 6:00 PM
 */
class EachCommandTest extends Specification{
    def "test init"(){
        when:
            def area = new BaseComand(new Pos(20, 20), new Size(4, 5))
            def eachCommand = new EachCommand(new Pos(1,2), new Size(2,3), "x", "dataList", area)
        then:
            assert eachCommand.pos == new Pos(1,2)
            assert eachCommand.initialSize == new Size(2,3)
            assert eachCommand.var == "x"
            assert eachCommand.items == "dataList"
            assert eachCommand.area == area
    }

    def "test size"(){
        given:
            def area = new BaseComand(new Pos(20, 20), new Size(4, 5))
            def eachCommand = new EachCommand(new Pos(1,2), new Size(2,3), "x", "dataList", area)
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
}
