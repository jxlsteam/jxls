package org.jxls.common;
import static org.junit.Assert.assertEquals;

import org.junit.Test;

/**
 * @author Leonid Vysochyn
 */
public class SizeTest {

    @Test
    public void construction() {
        Size size = new Size(20, 10);
        assertEquals( "Size height is incorrect", 10, size.getHeight());
        assertEquals( "Size width is incorrect", 20, size.getWidth());
    }

    @Test
    public void toStringTest() {
        Size size = new Size(20, 10);
        assertEquals("toString() result is incorrect", "(20,10)", size.toString());
    }

    @Test
    public void subtraction() {
        Size size1 = new Size(20, 10);
        Size size2 = new Size(10, 5);
        Size result = size1.minus(size2);
        assertEquals("Subtraction result is incorrect", new Size(10, 5), result);
    }

    @Test
    public void addition() {
        Size size1 = new Size(20, 10);
        Size size2 = new Size(10, 5);
        Size result = size1.add(size2);
        assertEquals("Sum is incorrect", new Size(30, 15), result);
    }
}
