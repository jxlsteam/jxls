package com.jxls.writer;

import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

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
        Pos origin = new Pos(5, 10);
        PosMap posMap1 = new PosMap(origin);
        posMap1.addToDestination( new Pos(10,10) );
        PosMap posMap2 = new PosMap(origin);
        posMap2.addToDestination(new Pos(15,10));
        posMap1.combine(posMap2);
        assertNotNull( posMap1 );
        PosMap posMap3 = new PosMap(origin);
        posMap3.addToDestination(new Pos(20,10));
        assertEquals("Combined by rows PosMap is incorrect", posMap3, posMap1);
    }

    @Test
    public void combinePosByCols(){
        Pos origin = new Pos(5, 10);
        PosMap posMap1 = new PosMap(origin);
        posMap1.addToDestination(new Pos(5, 15));
        PosMap posMap2 = new PosMap(origin);
        posMap2.addToDestination(new Pos(5, 20));
        posMap1.combine(posMap2);
        assertNotNull(posMap1);
        PosMap posMap3 = new PosMap(origin);
        posMap3.addToDestination(new Pos(5, 25));
        assertEquals("Combined by columns PosMap is incorrect", posMap3, posMap1);
    }

    @Test
    public void combineWithEmptyDestination(){
        Pos origin = new Pos(5, 10);
        PosMap posMap1 = new PosMap(origin);
        PosMap posMap2 = new PosMap(origin);
        posMap2.addToDestination(new Pos(10, 10) );
        posMap1.combine(posMap2);
        PosMap posMap3 = new PosMap(origin);
        posMap3.addToDestination(new Pos(10, 10));
        assertEquals("Combined PosMap is incorrect", posMap3, posMap1);
    }

}
