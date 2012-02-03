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
            1 == cellData.getSheetIndex()
            2 == cellData.getRowIndex()
            3 == cellData.getColIndex()
    }

}
