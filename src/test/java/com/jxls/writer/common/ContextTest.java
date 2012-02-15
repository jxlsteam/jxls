package com.jxls.writer.common;

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
        context.putVar("col", 1);
        context.putVar("row", 2);
        context.putVar("obj", "12345");
        Map vars = context.toMap();
        assertEquals(vars.get("col"), 1);
        assertEquals(vars.get("row"), 2);
        assertEquals(vars.get("obj"), "12345");
    }

    @Test
    public void removeVar(){
        Context context = new Context();
        context.putVar("col", 1);
        Map vars = context.toMap();
        assertEquals(vars.get("col"), 1);
        context.removeVar("col");
        vars = context.toMap();
        assertFalse( vars.containsKey("col") );
    }
    
    @Test 
    public void testToString(){
        Context context = new Context();
        context.putVar("x", 1);
        context.putVar("y", "Abc");
        assertEquals("Context{y=Abc, x=1}", context.toString() );
    }
}
