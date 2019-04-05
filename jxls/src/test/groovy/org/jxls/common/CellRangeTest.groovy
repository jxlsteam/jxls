package org.jxls.common

import spock.lang.Specification

/**
 * @author Leonid Vysochyn
 * Date: 1/26/12 2:56 PM
 */
class CellRangeTest extends Specification{
    def "test shift cells with row block"(){
        given:
            CellRange cellRange = new CellRange(new CellRef(0, 0), 10, 10)
            int startRow = 3;
            int endRow = 5;
            int col = 5;
            int shift = 3;
            cellRange.changeMatrix[5][6] = true;
        when:
            cellRange.shiftCellsWithRowBlock(startRow, endRow, col, shift, true)
        then:
            assert cellRange.getCell(4, 3) == new CellRef(4, 3)
            assert cellRange.getCell(3, 6) == new CellRef(3, 9)
            assert cellRange.getCell(5, 6) == new CellRef(5, 6)
            assert cellRange.getCell(6, 6) == new CellRef(6, 6)
            assert cellRange.getCell(4, 9) == new CellRef(4, 12)
    }

    def "test change matrix when shifting with row block"(){
        given:
            CellRange cellRange = new CellRange(new CellRef(0, 0), 10, 10)
            int startRow = 3;
            int endRow = 5;
            int col = 5;
            int shift = 3;
        when:
            cellRange.shiftCellsWithRowBlock(startRow, endRow, col, shift, true)
        then:
            assert cellRange.hasChanged(3, 6)
            assert cellRange.hasChanged(4, 7)
            assert !cellRange.hasChanged(6, 6)
            assert !cellRange.hasChanged(4, 2)
    }

    def "test shift cells with column block"(){
        given:
            CellRange cellRange = new CellRange(new CellRef(0, 0), 10, 10)
            int startCol = 3;
            int endCol = 5;
            int row = 5;
            int shift = -3;
            cellRange.changeMatrix[7][4] = true;
        when:
            cellRange.shiftCellsWithColBlock(startCol, endCol, row, shift, true)
        then:
            assert cellRange.getCell(3, 4) == new CellRef(3, 4)
            assert cellRange.getCell(6, 3) == new CellRef(3, 3)
            assert cellRange.getCell(7, 4) == new CellRef(7, 4)
            assert cellRange.getCell(7, 6) == new CellRef(7, 6)
            assert cellRange.getCell(7, 5) == new CellRef(4, 5)
    }

    def "test change matrix when shifting with column block"(){
        given:
            CellRange cellRange = new CellRange(new CellRef(0, 0), 10, 10)
            int startCol = 3;
            int endCol = 5;
            int row = 5;
            int shift = -3;
        when:
            cellRange.shiftCellsWithColBlock(startCol, endCol, row, shift, true)
        then:
            assert !cellRange.hasChanged(3, 4)
            assert !cellRange.hasChanged(7, 6)
            assert cellRange.hasChanged(6, 3)
            assert cellRange.hasChanged(7, 5)
    }

    def "test change matrix reset"(){
        given:
            CellRange cellRange = new CellRange(new CellRef(0, 0), 10, 10)
            int startCol = 3;
            int endCol = 5;
            int row = 5;
            int shift = -3;
            cellRange.shiftCellsWithColBlock(startCol, endCol, row, shift, true)
        when:
            cellRange.resetChangeMatrix()
        then:
            assert !cellRange.hasChanged(3, 4)
            assert !cellRange.hasChanged(7, 6)
            assert !cellRange.hasChanged(6, 3)
            assert !cellRange.hasChanged(7, 5)
    }

    def "test exclude cells"(){
        given:
            CellRange cellRange = new CellRange(new CellRef(0, 0), 10, 10)
        when:
            cellRange.excludeCells(3, 4, 5, 7)
        then:
            assert cellRange.isExcluded(5, 3)
            assert cellRange.isExcluded(6, 3)
            assert cellRange.isExcluded(7, 3)
            assert cellRange.isExcluded(5, 4)
            assert cellRange.isExcluded(6, 4)
            assert !cellRange.isExcluded(3, 3)
            assert !cellRange.isExcluded(5, 2)
            assert !cellRange.isExcluded(8, 4)
    }

}
