package com.jxls.writer

import spock.lang.Specification

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
            Util.createTargetCellRef([new CellRef("sheet1!A1"), new CellRef("sheet1!A5")]) == "sheet1!A1,sheet1!A5"
            Util.createTargetCellRef([new CellRef("sheet1!A1"), new CellRef("sheet1!A2"), new CellRef("sheet1!A3")]) == "sheet1!A1:sheet1!A3"
            Util.createTargetCellRef([new CellRef("sheet1!A1"), new CellRef("sheet1!A2"), new CellRef("sheet2!A3")]) == "sheet1!A1,sheet1!A2,sheet2!A3"
            Util.createTargetCellRef([new CellRef("sheet1!A1"), new CellRef("sheet1!A2"),
                    new CellRef("sheet1!A3"), new CellRef("sheet1!A5")]) == "sheet1!A1,sheet1!A2,sheet1!A3,sheet1!A5"
    }
    
    def "test group by col range"(){
        when:
            def posList = [new CellRef("sh!B10"), new CellRef("sh!C5"), new CellRef("sh!B2"), new CellRef("sh!A8"), new CellRef("sh!B4"), new CellRef("sh!C7"),
                    new CellRef("sh!B3"), new CellRef("sh!B11"), new CellRef("sh!D7"), new CellRef("sh!B1"), new CellRef("sh!E7")]
            def posList2 = [new CellRef("sh!B2"), new CellRef("sh!C2"), new CellRef("sh!D2"), new CellRef("sh!B5"), new CellRef("sh!B4"), new CellRef("sh!A2")]
        then:
            Util.groupByColRange(posList) == [
                    [new CellRef("sh!A8")],
                    [new CellRef("sh!B1"), new CellRef("sh!B2"), new CellRef("sh!B3"), new CellRef("sh!B4")],
                    [new CellRef("sh!B10"), new CellRef("sh!B11")],
                    [new CellRef("sh!C5")],
                    [new CellRef("sh!C7")],
                    [new CellRef("sh!D7")],
                    [new CellRef("sh!E7")]
            ]
            Util.groupByColRange(posList2) == [
                    [new CellRef("sh!A2")],
                    [new CellRef("sh!B2")],
                    [new CellRef("sh!B4"), new CellRef("sh!B5")],
                    [new CellRef("sh!C2")],
                    [new CellRef("sh!D2")]
            ]
    }

    def "test group by row range"(){
        when:
            def posList = [new CellRef("sh!B2"), new CellRef("sh!C2"), new CellRef("sh!D2"), new CellRef("sh!B5"), new CellRef("sh!B4"), new CellRef("sh!A2")]
        then:
            Util.groupByRowRange(posList) == [
                    [new CellRef("sh!A2"), new CellRef("sh!B2"), new CellRef("sh!C2"), new CellRef("sh!D2")],
                    [new CellRef("sh!B4")],
                    [new CellRef("sh!B5")]
            ]
    }

    def "test group by range with target range count setting"(){
        when:
            def posList = [new CellRef("sh!B2"), new CellRef("sh!C2"), new CellRef("sh!D2"), new CellRef("sh!B5"), new CellRef("sh!B4"), new CellRef("sh!A2")]
        then:
            Util.groupByRanges(posList, 3) == [
                    [new CellRef("sh!A2"), new CellRef("sh!B2"), new CellRef("sh!C2"), new CellRef("sh!D2")],
                    [new CellRef("sh!B4")],
                    [new CellRef("sh!B5")]
            ]
            Util.groupByRanges(posList, 5) == [
                    [new CellRef("sh!A2")],
                    [new CellRef("sh!B2")],
                    [new CellRef("sh!B4"), new CellRef("sh!B5")],
                    [new CellRef("sh!C2")],
                    [new CellRef("sh!D2")]
            ]

    }
}
