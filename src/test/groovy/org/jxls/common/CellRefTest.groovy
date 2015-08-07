package org.jxls.common

import spock.lang.Specification

/**
 * @author Leonid Vysochyn
 */
class CellRefTest extends Specification {

    def "test isValid"(){
        expect:
            new CellRef(-1, 5).isValid() == false
            new CellRef(1, -3).isValid() == false
            new CellRef(1, 0).isValid() == true
    }

    def "test create from cell ref"(){
        when:
            CellRef pos = new CellRef(cellref)
        then:
            pos == new CellRef(sheetName, row, col)
        where:
            cellref             | sheetName | row   | col
            "sheet1!A1"         | "sheet1"  | 0     | 0
            "sheet2!C5"         | "sheet2"  | 4     | 2
            "Sheet2!D5"         | "Sheet2"  | 4     | 3
    }
    
    def "test get cell name"(){
        expect:
            new CellRef("Sheet 3",3,3).getCellName() == "'Sheet 3'!D4"
    }

    def "test sorting"(){
        when:
            def posList = [new CellRef("B1"), new CellRef("B5"), new CellRef("B3"), new CellRef("C2"), new CellRef("D1"), new CellRef("A4")]
        then:
            posList.sort() == [new CellRef("A4"), new CellRef("B1"), new CellRef("B3"), new CellRef("B5"), new CellRef("C2"), new CellRef("D1") ]
    }
    
    def "test get formatted sheet name"(){
        expect:
            new CellRef(sheetName, 0, 0).getFormattedSheetName() == formattedSheetName
        where:
            sheetName   | formattedSheetName
            "Sheet1"    | "Sheet1"
            "Sheet 1"   | "'Sheet 1'"
            "1a"        | "'1a'"
    }
}
