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
        Cell cell = new Cell(10, 20);
        assertEquals("Cell col was set incorrectly", 10, cell.getCol() );
        assertEquals("Cell row was set incorrectly", 20, cell.getRow() );
    }

    @Test
    public void toStringTest(){
        Cell cell = new Cell(10, 20);
        assertEquals("toString() is incorrect", "(10,20)", cell.toString() );
    }
}
