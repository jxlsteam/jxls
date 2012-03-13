package com.jxls.writer.command

import spock.lang.Specification

import com.jxls.writer.common.Size
import com.jxls.writer.common.CellRef
import com.jxls.writer.area.XlsArea
import com.jxls.writer.area.Area
import com.jxls.writer.common.Context

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
    
    def "test create with CellRefGenerator"(){
        def cellRefGen = Mock(CellRefGenerator)
        def area = Mock(Area)
        when:
            def eachCommand = new EachCommand("item", "items", area, cellRefGen)
        then:
            eachCommand.var == "item"
            eachCommand.items == "items"
            eachCommand.area == area
            eachCommand.cellRefGenerator == cellRefGen
    }
    
    def "test add area"(){
        def area = Mock(Area)
        when:
            def command = new EachCommand("x", "list", EachCommand.Direction.DOWN)
            command.addArea(area)
        then:
            command.var == "x"
            command.items == "list"
            command.area == area
            command.areaList.size() == 1
    }
    
    def "test excessive number areas"(){
        def command = new EachCommand("x", "list", EachCommand.Direction.DOWN)
        command.addArea(Mock(Area))
        when:
            command.addArea(Mock(Area))
        then:
            thrown(IllegalArgumentException)
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
            eachArea.getSize() >> new Size(3, 2)
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

    def "test applyAt with CellRefGenerator"(){
        given:
        def area = Mock(Area)
        def cellRefGenerator = Mock(CellRefGenerator)
        def eachSheetCommand = new EachCommand("x", "list", area, cellRefGenerator)
        def context = Mock(Context)
        when:
        eachSheetCommand.applyAt(new CellRef("sheet2", 2,3), context)
        then:
        1 * context.toMap() >> ["list":[2,4,5]]
        1 * context.putVar("x", 2)
        1 * context.putVar("x", 4)
        1 * context.putVar("x", 5)
        1 * cellRefGenerator.generateCellRef(0, context) >> new CellRef("abc!A2")
        1 * cellRefGenerator.generateCellRef(1, context) >> new CellRef("def!B2")
        1 * cellRefGenerator.generateCellRef(2, context) >> new CellRef("ghi!C2")
        1 * area.applyAt(new CellRef("abc!A2"), context) >> new Size(3,5)
        1 * area.applyAt(new CellRef("def!B2"), context) >> new Size(2,3)
        1 * area.applyAt(new CellRef("ghi!C2"), context) >> new Size(4,3)
    }

    def "test select attribute"(){
        def eachArea = Mock(Area)
        def context = new Context()
        context.putVar("items", [0,1,2,3,4,5,1])
        def eachCommand = new EachCommand("var", "items", eachArea)
        when:
            eachCommand.setSelect(" var % 2 ==0")
            eachCommand.applyAt(new CellRef("sheet1!A1"), context)
        then:
            3 * eachArea.applyAt(_, context ) >> new Size(1,2)
            0 * _
    }
}
