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
        Pos pos = new Pos(10, 20);
        assertEquals("Pos row was set incorrectly", 10, pos.getRow() );
        assertEquals("Pos col was set incorrectly", 20, pos.getCol() );
    }

    @Test
    public void toStringTest(){
        Pos pos = new Pos(10, 20);
        assertEquals("toString() is incorrect", "(10,20)", pos.toString() );
    }
}
