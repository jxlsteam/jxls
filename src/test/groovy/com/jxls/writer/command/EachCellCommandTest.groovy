package com.jxls.writer.command

import spock.lang.Specification
import com.jxls.writer.area.Area
import com.jxls.writer.common.CellRef
import com.jxls.writer.common.Context

/**
 * @author Leonid Vysochyn
 * Date: 2/17/12 5:15 PM
 */
class EachCellCommandTest extends Specification{
    def "test applyAt"(){
        given:
            def area = Mock(Area)
            def cellRefGenerator = Mock(CellRefGenerator)
            def eachSheetCommand = new EachCellCommand("x", "list", cellRefGenerator, area)
            def context = Mock(Context)
        when:
            eachSheetCommand.applyAt(new CellRef("sheet2", 2,3), context)
        then:
            1 * cellRefGenerator.generateCellRef(0, context) >> new CellRef("abc!A2")
            1 * cellRefGenerator.generateCellRef(1, context) >> new CellRef("def!B2")
            1 * cellRefGenerator.generateCellRef(2, context) >> new CellRef("ghi!C2")
            1 * area.applyAt(new CellRef("abc!A2"))
    }
}
