package org.jxls.command

import spock.lang.Specification
import org.jxls.area.Area
import org.jxls.common.Size

/**
 * @author Leonid Vysochyn
 * Date: 6/19/12 4:56 PM
 */
class AbstractCommandTest extends Specification{
    def "test reset"(){
        setup:
            def area1 = Mock(Area)
            def area2 = Mock(Area)
            def area3 = Mock(Area)
            def command = [name: {->'testcommand'}, applyAt: {-> new Size(0,0)} ] as AbstractCommand
            command.addArea(area1)
            command.addArea(area2)
            command.addArea(area3)
        when:
            command.reset()
        then:
            1 * area1.reset()
            1 * area2.reset()
            1 * area3.reset()
            0 * _._
    }
}
