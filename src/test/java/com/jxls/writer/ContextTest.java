package com.jxls.writer;

import com.jxls.writer.command.Context;
import org.junit.Test;

import java.util.Map;

import static junit.framework.Assert.assertEquals;
import static junit.framework.Assert.assertFalse;

/**
 * Date: Nov 10, 2009
 *
 * @author Leonid Vysochyn
 */
public class ContextTest {
    @Test
    public void addVars(){
        Context context = new Context();
        context.putVar("x", 1);
        context.putVar("y", 2);
        context.putVar("obj", "12345");
        Map vars = context.toMap();
        assertEquals(vars.get("x"), 1);
        assertEquals(vars.get("y"), 2);
        assertEquals(vars.get("obj"), "12345");
    }

    @Test
    public void removeVar(){
        Context context = new Context();
        context.putVar("x", 1);
        Map vars = context.toMap();
        assertEquals(vars.get("x"), 1);
        context.removeVar("x");
        vars = context.toMap();
        assertFalse( vars.containsKey("x") );
    }
}
