package org.jxls.command

import org.jxls.area.Area
import org.jxls.area.XlsArea
import org.jxls.common.CellRef
import org.jxls.common.Context
import org.jxls.common.Size
import spock.lang.Specification

/**
 * Created by Leonid Vysochyn on 28-Jun-15.
 */
class GridCommandTest extends Specification{
    def "test init"(){
        when:
        def headerArea = new XlsArea(new CellRef(10, 10), new Size(1, 1))
        def bodyArea = new XlsArea(new CellRef(10, 11), new Size(1, 1))
        def gridCommand = new GridCommand( "hs", "datas", headerArea, bodyArea)
        then:
        gridCommand.headers == "hs"
        gridCommand.data == "datas"
        gridCommand.headerArea == headerArea
        gridCommand.bodyArea == bodyArea
    }

    def "test add area"(){
        def headerArea = Mock(Area)
        def bodyArea = Mock(Area)
        when:
        def command = new GridCommand("hs", "list")
        command.addArea(headerArea)
        command.addArea(bodyArea)
        then:
        command.headers == "hs"
        command.data == "list"
        command.headerArea == headerArea
        command.bodyArea == bodyArea
        command.areaList.size() == 2
    }

    def "test excessive number areas"(){
        def command = new GridCommand("hs", "list")
        command.addArea(Mock(Area))
        command.addArea(Mock(Area))
        when:
        command.addArea(Mock(Area))
        then:
        thrown(IllegalArgumentException)
    }

    def "test applyAt with row collection"(){
        given:
        def headerArea = Mock(Area)
        def bodyArea = Mock(Area)
        def gridCommand = new GridCommand( "hs", "datas", headerArea, bodyArea)
        def context = Mock(Context)
        when:
        gridCommand.applyAt(new CellRef("sheet2", 2, 2), context)
        then:
        context.toMap() >> ["hs": ["H1","H2","H3"], "datas": [[10, 11, 12], [20,21,22],[30,31,32]]]
        1 * context.putVar(GridCommand.HEADER_VAR, "H1")
        1 * context.putVar(GridCommand.HEADER_VAR, "H2")
        1 * context.putVar(GridCommand.HEADER_VAR, "H3")
        1 * context.removeVar(GridCommand.HEADER_VAR)
        1 * headerArea.applyAt(new CellRef("sheet2", 2, 2), context) >> new Size(1, 2)
        1 * headerArea.applyAt(new CellRef("sheet2", 2, 3), context) >> new Size(2, 3)
        1 * headerArea.applyAt(new CellRef("sheet2", 2, 5), context) >> new Size(2, 4)

        1 * context.putVar(GridCommand.DATA_VAR, 10)
        1 * context.putVar(GridCommand.DATA_VAR, 11)
        1 * context.putVar(GridCommand.DATA_VAR, 12)
        1 * context.putVar(GridCommand.DATA_VAR, 20)
        1 * context.putVar(GridCommand.DATA_VAR, 21)
        1 * context.putVar(GridCommand.DATA_VAR, 22)
        1 * context.putVar(GridCommand.DATA_VAR, 30)
        1 * context.putVar(GridCommand.DATA_VAR, 31)
        1 * context.putVar(GridCommand.DATA_VAR, 32)
        1 * context.removeVar(GridCommand.DATA_VAR)

        1 * bodyArea.applyAt(new CellRef("sheet2", 6, 2), context) >> new Size(2, 1)
        1 * bodyArea.applyAt(new CellRef("sheet2", 6, 4), context) >> new Size(2, 3)
        1 * bodyArea.applyAt(new CellRef("sheet2", 6, 6), context) >> new Size(2, 1)
        1 * bodyArea.applyAt(new CellRef("sheet2", 9, 2), context) >> new Size(1, 2)
        1 * bodyArea.applyAt(new CellRef("sheet2", 9, 3), context) >> new Size(1, 1)
        1 * bodyArea.applyAt(new CellRef("sheet2", 9, 4), context) >> new Size(1, 1)
        1 * bodyArea.applyAt(new CellRef("sheet2", 11, 2), context) >> new Size(2, 4)
        1 * bodyArea.applyAt(new CellRef("sheet2", 11, 4), context) >> new Size(1, 3)
        1 * bodyArea.applyAt(new CellRef("sheet2", 11, 5), context) >> new Size(1, 1)
        0 * _._
    }
}
