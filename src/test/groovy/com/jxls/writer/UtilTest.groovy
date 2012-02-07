package com.jxls.writer

import spock.lang.Specification

/**
 * @author Leonid Vysochyn
 */
class UtilTest extends Specification{
    def "test get formula cell refs"(){
        when:
            def cellRefs = Util.getFormulaCellRefs('SUM(A1:A10) + AE100*BC12 + Sheet2!B12 + 5')
        then:
            cellRefs.toArray() == ["A1", "A10", "AE100", "BC12", "Sheet2!B12"]
    }

    def "test get formula cell refs with union cells"(){
        when:
            def cellRefs = Util.getFormulaCellRefs('SUM(A1:A10) + AE100*BC12 + Sheet2!B12 + 5 + U(A1,A2)')
        then:
            cellRefs.toArray() == ["A1", "A10", "AE100", "BC12", "Sheet2!B12"]
    }

    def "test create target cell ref"(){
        expect:
            Util.createTargetCellRef([new Pos("A1"), new Pos("A5")]) == "A1,A5"
    }
    
    def "test group by ranges"(){
        when:
            def posList = [new Pos("B10"), new Pos("C5"), new Pos("B2"), new Pos("A8"), new Pos("B4"), new Pos("C7"),
                    new Pos("B3"), new Pos("B11"), new Pos("D7"), new Pos("B1"), new Pos("E7")]
        then:
            Util.groupByRanges(posList) == [
                    [new Pos("B1"), new Pos("B2"), new Pos("B3"), new Pos("B4")],
                    [new Pos("C5")],
                    [new Pos("C7"), new Pos("D7"), new Pos("E7")],
                    [new Pos("A8")],
                    [new Pos("B10"), new Pos("B11")]
            ]
    }
}
