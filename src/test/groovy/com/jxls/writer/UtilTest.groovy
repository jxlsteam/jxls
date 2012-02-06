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

    def "test create target cell ref"(){
        expect:
            Util.createTargetCellRef([new Pos("A1"), new Pos("A5")]) == "A1,A5"
    }
}
