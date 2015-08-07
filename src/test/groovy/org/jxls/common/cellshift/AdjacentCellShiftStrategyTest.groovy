package org.jxls.common.cellshift

import org.jxls.common.CellRef
import spock.lang.Specification

/**
 * Created by Leonid Vysochyn on 07-Aug-15.
 */
class AdjacentCellShiftStrategyTest extends Specification{
    def "test requiresColShifting"(){
        given:
        CellShiftStrategy strategy = new AdjacentCellShiftStrategy()
        CellRef cellRef = new CellRef(5, 3)
        expect:
        strategy.requiresColShifting(cellRef, 3, 6, 2) == true
        strategy.requiresColShifting(cellRef, 3, 6, 3) == false
        strategy.requiresColShifting(cellRef, 1, 4, 2) == true
        strategy.requiresColShifting(cellRef, 6, 7, 2) == true
        strategy.requiresColShifting(cellRef, 6, 7, 3) == false
    }

    def "test requiresRowShifting"(){
        given:
        CellShiftStrategy strategy = new AdjacentCellShiftStrategy()
        CellRef cellRef = new CellRef(5, 6)
        expect:
        strategy.requiresRowShifting(cellRef, 4, 7, 3) == true
        strategy.requiresRowShifting(cellRef, 7, 8, 3) == true
        strategy.requiresRowShifting(cellRef, 2, 5, 2) == true
        strategy.requiresRowShifting(cellRef, 2, 6, 4) == true
        strategy.requiresRowShifting(cellRef, 2, 6, 5) == false
    }
}
