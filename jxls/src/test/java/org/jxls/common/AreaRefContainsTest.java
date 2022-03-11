package org.jxls.common;

import org.junit.Assert;
import org.junit.Test;

/**
 * AreaRef.contains() testcases
 */
public class AreaRefContainsTest {
    private static final AreaRef areaRef = new AreaRef(new CellRef(4, 4), new CellRef(6, 6));

    @Test
    public void above() {
        Assert.assertFalse(areaRef.contains(new CellRef(3, 5)));
    }
    
    @Test
    public void top() {
        Assert.assertTrue(areaRef.contains(new CellRef(4, 5)));
    }
    
    @Test
    public void bottom() {
        Assert.assertTrue(areaRef.contains(new CellRef(6, 5)));
    }
    
    @Test
    public void below() {
        Assert.assertFalse(areaRef.contains(new CellRef(7, 5)));
    }
    
    @Test
    public void outsideLeft() {
        Assert.assertFalse(areaRef.contains(new CellRef(5, 3)));
    }
    
    @Test
    public void insideLeft() {
        Assert.assertTrue(areaRef.contains(new CellRef(5, 4)));
    }
    
    @Test
    public void insideRight() {
        Assert.assertTrue(areaRef.contains(new CellRef(5, 6)));
    }
    
    @Test
    public void outsideRight() {
        Assert.assertFalse(areaRef.contains(new CellRef(5, 7)));
    }
    
    @Test
    public void sheetName() {
        AreaRef s = new AreaRef(new CellRef("abc", 4, 4), new CellRef("abc", 6, 6));
        Assert.assertTrue(s.contains(new CellRef("abc", 5, 5)));
    }
    
    @Test
    public void otherSheetName() {
        AreaRef s = new AreaRef(new CellRef("abc", 4, 4), new CellRef("abc", 6, 6));
        Assert.assertFalse(s.contains(new CellRef("def", 5, 5)));
        Assert.assertFalse(s.contains(new CellRef(5, 5)));
    }
}
