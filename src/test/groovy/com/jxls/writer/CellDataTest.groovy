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

}
