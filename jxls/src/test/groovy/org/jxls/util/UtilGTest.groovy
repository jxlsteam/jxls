package org.jxls.util

import org.jxls.expression.JexlExpressionEvaluator
import spock.lang.Specification
import org.jxls.common.CellRef

import org.jxls.common.Context
import org.jxls.expression.Dummy

/**
 * @author Leonid Vysochyn
 */
class UtilGTest extends Specification{
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

    def "test boolean condition calculation"(){
        given:
        def context = new Context()
        def evaluator = new JexlExpressionEvaluator()
        when:
            context.putVar("x", xValue)
        then:
            Util.isConditionTrue(evaluator, "2*x + 5 > 10", context) == result
        where:
            xValue  | result
            2       | false
            3       | true
    }

    def "test get object property"(){
        expect:
            Util.getObjectProperty(new Dummy(str, val), prop) == result
        where:
            str  |  val | prop      | result
            "Abc"| 0    | "strValue"| "Abc"
            "x y"| 0    | "strValue"| "x y"
            null | 0    | "strValue"| null
            null | -5   | "intValue"| -5
            ""   | 100  | "intValue"| 100
    }
    
    def "test group collection"(){
        def list = [new Dummy("abc", 1), new Dummy("ab", 2), new Dummy("abc", 3), new Dummy("bc", 2), new Dummy("ac", 4), new Dummy("ab", 1)]
        when:
            def groups = Util.groupCollection( list, "strValue", "asc" )
        then:
            groups.size() == 4
            groups[0].item == new Dummy("ab", 2)
            groups[0].items == [ new Dummy("ab", 2), new Dummy("ab", 1) ]
            groups[1].item == new Dummy("abc", 1)
            groups[1].items == [ new Dummy("abc", 1), new Dummy("abc", 3)]
            groups[2].item == new Dummy("ac", 4)
            groups[2].items == [ new Dummy("ac", 4) ]
            groups[3].item == new Dummy("bc", 2)
            groups[3].items == [new Dummy("bc", 2)]
    }

    def "test group collection with null keys"() {
        def list = [new Dummy("abc", 1), new Dummy(null, 4), new Dummy("abc", 3), new Dummy(null, 2)]
        when:
            def groups = Util.groupCollection(list, "strValue", "asc")
        then:
            groups.size() == 2
            groups[0].item == new Dummy("abc", 1)
            groups[0].items == [new Dummy("abc", 1), new Dummy("abc", 3)]
            groups[1].item == new Dummy(null, 4)
            groups[1].items == [new Dummy(null, 4), new Dummy(null, 2)]
    }
}
