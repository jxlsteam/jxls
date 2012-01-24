package com.jxls.writer;

import org.junit.Test;

import static junit.framework.Assert.assertEquals;
import static junit.framework.Assert.assertNotNull;

/**
 * Date: Nov 3, 2009
 *
 * @author Leonid Vysochyn
 */
public class PosMapTest {

    @Test
    public void combinePosByRows(){
        Cell origin = new Cell(5, 10);
        PosMap posMap1 = new PosMap(origin);
        posMap1.addToDestination( new Cell(10,10) );
        PosMap posMap2 = new PosMap(origin);
        posMap2.addToDestination(new Cell(15,10));
        posMap1.combine(posMap2);
        assertNotNull( posMap1 );
        PosMap posMap3 = new PosMap(origin);
        posMap3.addToDestination(new Cell(20,10));
        assertEquals("Combined by rows PosMap is incorrect", posMap3, posMap1);
    }

    @Test
    public void combinePosByCols(){
        Cell origin = new Cell(5, 10);
        PosMap posMap1 = new PosMap(origin);
        posMap1.addToDestination(new Cell(5, 15));
        PosMap posMap2 = new PosMap(origin);
        posMap2.addToDestination(new Cell(5, 20));
        posMap1.combine(posMap2);
        assertNotNull(posMap1);
        PosMap posMap3 = new PosMap(origin);
        posMap3.addToDestination(new Cell(5, 25));
        assertEquals("Combined by columns PosMap is incorrect", posMap3, posMap1);
    }

    @Test
    public void combineWithEmptyDestination(){
        Cell origin = new Cell(5, 10);
        PosMap posMap1 = new PosMap(origin);
        PosMap posMap2 = new PosMap(origin);
        posMap2.addToDestination(new Cell(10, 10) );
        posMap1.combine(posMap2);
        PosMap posMap3 = new PosMap(origin);
        posMap3.addToDestination(new Cell(10, 10));
        assertEquals("Combined PosMap is incorrect", posMap3, posMap1);
    }

}
