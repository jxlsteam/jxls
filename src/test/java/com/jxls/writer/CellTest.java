package com.jxls.writer;

import org.junit.Test;
import static org.junit.Assert.assertEquals;

/**
 * Date: Mar 13, 2009
 *
 * @author Leonid Vysochyn
 */
public class CellTest {
    @Test
    public void construction(){
        Cell cell = new Cell(1, 20, 10);
        assertEquals("Cell sheet was set incorrectly", 1, cell.getSheetIndex());
        assertEquals("Cell row was set incorrectly", 20, cell.getRow() );
        assertEquals("Cell col was set incorrectly", 10, cell.getCol() );
    }

    @Test
    public void toStringTest(){
        Cell cell = new Cell(20, 10);
        assertEquals("toString() is incorrect", "Cell(0,20,10)", cell.toString() );
    }
}
