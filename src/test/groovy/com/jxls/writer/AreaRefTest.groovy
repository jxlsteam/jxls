package com.jxls.writer

import spock.lang.Specification

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
}
