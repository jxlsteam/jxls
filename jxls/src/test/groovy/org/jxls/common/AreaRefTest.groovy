package org.jxls.common

import spock.lang.Specification

/**
 * @author Leonid Vysochyn
 * @since 2/13/2012
 */
class AreaRefTest extends Specification{

    def "test create with first & last cell refs"() {
        when:
            def areaRef = new AreaRef(new CellRef("sheet1!A1"), new CellRef("sheet1!D4"))
        then:
            areaRef.getFirstCellRef() == new CellRef("sheet1", 0, 0)
            areaRef.getLastCellRef() == new CellRef("sheet1", 3, 3)
    }
    
    def "test create with incorrect cell refs"() {
        when:
            def areaRef = new AreaRef(new CellRef("sheet1!A1"), new CellRef("sheet2!D4"))
        then:
            def e = thrown(IllegalArgumentException)
            e.cause == null
            e.message == "Cannot create area from specified cell references sheet1!A1, sheet2!D4"
    }

    def "test create with area ref"() {
        when:
            def areaRef = new AreaRef("B3:D5")
        then:
            areaRef.getFirstCellRef() == new CellRef("B3")
            areaRef.getLastCellRef() == new CellRef("D5")
    }

    def "test get area size"() {
        when:
            def areaRef = new AreaRef("A1:C4")
        then:
            areaRef.getSize() == new Size(3,4)
    }
    
    def "test create with cellRef and size"() {
        when:
            def areaRef = new AreaRef(new CellRef("sheet1!B2"), new Size(1,2))
        then:
            areaRef.firstCellRef == new CellRef("sheet1!B2")
            areaRef.getSize() == new Size(1,2)
            areaRef.lastCellRef == new CellRef("sheet1!B3")
    }
    
    def "test get sheet name"() {
        when:
            def areaRef = new AreaRef(new CellRef("sheet1!B2"), new Size(1,1))
        then:
            areaRef.sheetName == "sheet1"
    }

    def "test contains another area"() {
        def area = new AreaRef("sheet1!B1:G8")
        expect:    area.contains(new AreaRef(areaRef)) == result
        where:      areaRef         | result
                    "sheet1!A2:B2"  | false
                    "sheet1!F2:H3"  | false
                    "sheet1!E4:G5"  | true
                    "sheet2!E4:G5"  | false
                    "sheet1!D8:D9"  | false
                    "sheet1!C8:D8"  | true
    }
    
    def "test toString"() {
        expect:
            new AreaRef(areaRef).toString() == result
        where:
            areaRef             | result
            "sheet1!A2:B2"      | "sheet1!A2:B2"
            "'sheet 1'!F2:H3"   | "'sheet 1'!F2:H3"
    }
}
