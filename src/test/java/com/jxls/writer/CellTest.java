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
        assertEquals("Pos x was set incorrectly", 10, pos.getX() );
        assertEquals("Pos y was set incorrectly", 20, pos.getY() );
    }

    @Test
    public void toStringTest(){
        Pos pos = new Pos(10, 20);
        assertEquals("toString() is incorrect", "(10,20)", pos.toString() );
    }
}
