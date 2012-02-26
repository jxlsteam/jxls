package com.jxls.writer.common

import spock.lang.Specification
import com.jxls.writer.common.CellData
import com.jxls.writer.common.CellRef

/**
 * @author Leonid Vysochyn
 * Date: 2/3/12 1:19 PM
 */
class CellDataTest extends Specification{
    def "test creation with (sheetName, col, row) params"(){
        when:
            CellData cellData = new CellData("sheet1", 2, 3)
        then:
            "sheet1" == cellData.getSheetName()
            2 == cellData.getRow()
            3 == cellData.getCol()
    }

    def "test equality"(){
        expect:
            new CellData("sheet1", 5, 10, CellData.CellType.STRING, "Abc").equals( new CellData("sheet1", 5, 10, CellData.CellType.STRING, "Abc"))
            !new CellData("sheet1", 5, 10, CellData.CellType.STRING, "Abc1").equals( new CellData("sheet1", 5, 10, CellData.CellType.STRING, "Abc"))
    }
    
    def "test creation with (sheetName, row, col, type, value)"(){
        when:
            CellData cellData = new CellData("Sheet1", 5, 10, CellData.CellType.NUMBER, 1.2)
        then:
            "Sheet1" == cellData.getSheetName()
            5 == cellData.getRow()
            10 == cellData.getCol()
            CellData.CellType.NUMBER == cellData.getCellType()
            1.2 == cellData.getCellValue()
            new CellRef("Sheet1",5,10) == cellData.getCellRef()
    }
    
    def "test creation with (pos, type, value)"(){
        when: 
            CellData cellData = new CellData(new CellRef("Sheet1",5,10), CellData.CellType.STRING, "Abc")
        then:
            new CellRef("Sheet1",5,10) == cellData.getCellRef()
            "Sheet1" == cellData.getSheetName()
            5 == cellData.getRow()
            10 == cellData.getCol()
    }


    def "test add target pos"(){
        when:
            CellData cellData = new CellData("sheet1", 2, 3)
            cellData.addTargetPos(new CellRef("sheet1",2,3))
            cellData.addTargetPos(new CellRef("sheet2",3,4))
            def targetPos = cellData.getTargetPos()
        then:
            targetPos.size() == 2
            targetPos.contains(new CellRef("sheet1", 2,3))
            targetPos.contains(new CellRef("sheet2",3,4))
    }    
    
    def "test get pos"(){
        when:
            CellData cellData = new CellData("sheet1",2,3)
        then:
            cellData.getCellRef() == new CellRef("sheet1",2,3)
    }
    
    def "test create with formula value"(){
        expect:
            new CellData("sheet1", 1, 2, cellType, formula).getFormula() == formulaValue
        where:
            formula             | formulaValue      | cellType
            "SUM(A1:A3)+B10"    | "SUM(A1:A3)+B10"  | CellData.CellType.FORMULA
            '$[SUM(A3:A7)+B12]' | "SUM(A3:A7)+B12"  | CellData.CellType.STRING
    }
    
    def "test reset target pos"(){
        when:
            def cellData = new CellData("sheet1",1,2, CellData.CellType.STRING, "Abc")
            cellData.addTargetPos(new CellRef("sheet2",0,0))
            cellData.addTargetPos(new CellRef("sheet1",1,1))
            cellData.resetTargetPos()
        then:
            cellData.getTargetPos().isEmpty()
    }

    def "test set/get comment"(){
        when:
            def cellData = new CellData("sheet1", 1,2, CellData.CellType.STRING, "Abc")
            cellData.setCellComment("Test comment")
        then:
            cellData.getCellComment() == "Test comment"
    }

}
