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
            pos == new Pos(sheet, row, col)
        where:
            cellref | sheet | row   | col
            "A1"    | 0     | 0     | 0
            "C5"    | 0     | 4     | 2
    }
    
    def "test get cell name"(){
        expect:
            new Pos(0,3,3).getCellName() == "D4"
    }
}
