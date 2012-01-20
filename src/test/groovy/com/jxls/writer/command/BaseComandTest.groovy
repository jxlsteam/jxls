package com.jxls.writer.command

import spock.lang.Specification
import com.jxls.writer.Pos
import com.jxls.writer.Size

/**
 * @author Leonid Vysochyn
 * Date: 1/18/12 6:25 PM
 */
class BaseComandTest extends Specification{
    def "test init"(){
        when:
            def area = new BaseComand(new Pos(1,1), new Size(5,5))
        then:
            assert area.pos == new Pos(1,1)
    }
}
