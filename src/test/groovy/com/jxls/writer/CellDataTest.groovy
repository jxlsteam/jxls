package com.jxls.writer

import spock.lang.Specification

/**
 * @author Leonid Vysochyn
 * Date: 2/3/12 1:19 PM
 */
class CellDataTest extends Specification{
    def "test creation with (sheet, col, row) params"(){
        when:
            CellData cellData = new CellData(1, 2, 3)
        then:
            1 == cellData.getSheet()
            2 == cellData.getRow()
            3 == cellData.getCol()
    }
    
    def "test creation with (sheet, col, row, type, value)"(){
        when:
            CellData cellData = new CellData(0, 5, 10, CellData.CellType.NUMBER, 1.2)
        then:
            0 == cellData.getSheet()
            5 == cellData.getRow()
            10 == cellData.getCol()
            CellData.CellType.NUMBER == cellData.getCellType()
            1.2 == cellData.getCellValue()
    }


    def "test add target pos"(){
        when:
            CellData cellData = new CellData(1, 2, 3)
            cellData.addTargetPos(new Pos(2,3))
            cellData.addTargetPos(new Pos(1,3,4))
            def targetPos = cellData.getTargetPos()
        then:
            targetPos.size() == 2
            targetPos.contains(new Pos(2,3))
            targetPos.contains(new Pos(1,3,4))
    }    
    
    def "test get pos"(){
        when:
            CellData cellData = new CellData(1,2,3)
        then:
            cellData.getPos() == new Pos(1,2,3)
    }
    
    def "test create with formula value"(){
        expect:
            new CellData(0, 1, 2, cellType, formula).getFormula() == formulaValue
        where:
            formula             | formulaValue      | cellType
            "SUM(A1:A3)+B10"    | "SUM(A1:A3)+B10"  | CellData.CellType.FORMULA
            '$[SUM(A3:A7)+B12]' | "SUM(A3:A7)+B12"  | CellData.CellType.STRING
    }

}
