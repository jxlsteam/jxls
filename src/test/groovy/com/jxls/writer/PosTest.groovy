package com.jxls.writer

import spock.lang.Specification

/**
 * @author Leonid Vysochyn
 */
class PosTest extends Specification {
    def "test create from cell ref"(){
        when:
            Pos pos = new Pos(cellref)
        then:
            pos == new Pos(row, col)
        where:
            cellref     | sheetName | row   | col
            "A1"        | null      | 0     | 0
            "C5"        | null      | 4     | 2
            "Sheet2!D5" | null      | 4     | 3
    }
    
    def "test get cell name"(){
        expect:
            new Pos("Sheet 3",3,3).getCellName() == "'Sheet 3'!D4"
    }

    def "test sorting"(){
        when:
            def posList = [new Pos("B1"), new Pos("B5"), new Pos("B3"), new Pos("C2"), new Pos("D1"), new Pos("A4")]
        then:
            posList.sort() == [new Pos("A4"), new Pos("B1"), new Pos("B3"), new Pos("B5"), new Pos("C2"), new Pos("D1") ]
    }
}
