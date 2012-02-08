package com.jxls.writer

import spock.lang.Specification
import spock.lang.Ignore

/**
 * @author Leonid Vysochyn
 */
class UtilTest extends Specification{
    def "test get formula cell refs"(){
        when:
            def cellRefs = Util.getFormulaCellRefs('SUM(A1:A10) + AE100*BC12 + Sheet2!B12 + 5 + U_(A1,A8)')
        then:
            cellRefs.toArray() == ["A1", "A10", "AE100", "BC12", "Sheet2!B12"]
    }

    def "test get jointed cell refs"(){
        when:
            def cellRefs = Util.getJointedCellRefs('SUM(A1:A10) + AE100*BC12 + 5 + U_(A1,A2) + B2 + U_(B1,B2,B3)')
        then:
            cellRefs.toArray() == ["U_(A1,A2)", "U_(B1,B2,B3)"]
    }
    
    def "test get cell refs from jointed cell ref"(){
        when:
            def cellRefs = Util.getCellRefsFromJointedCellRef('U_(B1,B4,C9,C10)')
        then:
            cellRefs.toArray() == ["B1","B4","C9","C10"]
    }
    
    def "test string contains jointed cell ref"(){
        expect:
            Util.formulaContainsJointedCellRef(formula) == result
        where:
            formula                 | result
            "SUM(U_(U1,U2)) + B3"   | true
            "SUM(U1,U2) + B3"       | false
    }

    def "test create target cell ref"(){
        expect:
            Util.createTargetCellRef([new Pos("A1"), new Pos("A5")]) == "A1,A5"
            Util.createTargetCellRef([new Pos("A1"), new Pos("A2"), new Pos("A3")]) == "A1:A3"
            Util.createTargetCellRef([new Pos("A1"), new Pos("A2"), new Pos("A3"), new Pos("A5")]) == "A1,A2,A3,A5"
    }
    
    def "test group by col range"(){
        when:
            def posList = [new Pos("B10"), new Pos("C5"), new Pos("B2"), new Pos("A8"), new Pos("B4"), new Pos("C7"),
                    new Pos("B3"), new Pos("B11"), new Pos("D7"), new Pos("B1"), new Pos("E7")]
            def posList2 = [new Pos("B2"), new Pos("C2"), new Pos("D2"), new Pos("B5"), new Pos("B4"), new Pos("A2")]
        then:
            Util.groupByColRange(posList) == [
                    [new Pos("A8")],
                    [new Pos("B1"), new Pos("B2"), new Pos("B3"), new Pos("B4")],
                    [new Pos("B10"), new Pos("B11")],
                    [new Pos("C5")],
                    [new Pos("C7")],
                    [new Pos("D7")],
                    [new Pos("E7")]
            ]
            Util.groupByColRange(posList2) == [
                    [new Pos("A2")],
                    [new Pos("B2")],
                    [new Pos("B4"), new Pos("B5")],
                    [new Pos("C2")],
                    [new Pos("D2")]
            ]
    }

    def "test group by row range"(){
        when:
            def posList = [new Pos("B2"), new Pos("C2"), new Pos("D2"), new Pos("B5"), new Pos("B4"), new Pos("A2")]
        then:
            Util.groupByRowRange(posList) == [
                    [new Pos("A2"), new Pos("B2"), new Pos("C2"), new Pos("D2")],
                    [new Pos("B4")],
                    [new Pos("B5")]
            ]
    }

    def "test group by range with target range count setting"(){
        when:
            def posList = [new Pos("B2"), new Pos("C2"), new Pos("D2"), new Pos("B5"), new Pos("B4"), new Pos("A2")]
        then:
            Util.groupByRanges(posList, 3) == [
                    [new Pos("A2"), new Pos("B2"), new Pos("C2"), new Pos("D2")],
                    [new Pos("B4")],
                    [new Pos("B5")]
            ]
            Util.groupByRanges(posList, 5) == [
                    [new Pos("A2")],
                    [new Pos("B2")],
                    [new Pos("B4"), new Pos("B5")],
                    [new Pos("C2")],
                    [new Pos("D2")]
            ]

    }
}
