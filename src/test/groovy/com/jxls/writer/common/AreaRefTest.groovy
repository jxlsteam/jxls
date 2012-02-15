package com.jxls.writer.common

import spock.lang.Specification
import com.jxls.writer.common.AreaRef
import com.jxls.writer.common.CellRef
import com.jxls.writer.common.Size

/**
 * @author Leonid Vysochyn
 * Date: 2/13/12 1:31 PM
 */
class AreaRefTest extends Specification{
    def "test create with first & last cell refs"(){
        when:
            def areaRef = new AreaRef(new CellRef("sheet1!A1"), new CellRef("sheet1!D4"))
        then:
            areaRef.getFirstCellRef() == new CellRef("sheet1", 0, 0)
            areaRef.getLastCellRef() == new CellRef("sheet1", 3, 3)
    }

    def "test create with area ref"(){
        when:
            def areaRef = new AreaRef("B3:D5")
        then:
            areaRef.getFirstCellRef() == new CellRef("B3")
            areaRef.getLastCellRef() == new CellRef("D5")
    }

    def "test get area size"(){
        when:
            def areaRef = new AreaRef("A1:C4")
        then:
            areaRef.getSize() == new Size(3,4)
    }
}
