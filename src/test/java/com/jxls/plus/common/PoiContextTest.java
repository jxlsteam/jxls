package com.jxls.plus.common;

import com.jxls.plus.common.Context;
import com.jxls.plus.transform.poi.PoiContext;
import com.jxls.plus.transform.poi.PoiUtil;
import org.junit.Test;

import java.util.HashMap;
import java.util.Map;

import static junit.framework.Assert.assertEquals;
import static junit.framework.Assert.assertNotNull;
import static junit.framework.Assert.assertTrue;

/**
 * @author Leonid Vysochyn
 *         Date: 6/18/12 5:13 PM
 */
public class PoiContextTest {
    @Test
    public void createEmptyContext(){
        Context context = new PoiContext();
        Object poiObject = context.getVar(PoiContext.POI_OBJECT_KEY);
        assertNotNull(poiObject);
        assertTrue(poiObject instanceof PoiUtil);
    }

    @Test
    public void createNonEmptyContext(){
        Map vars = new HashMap();
        vars.put("a", 123);
        Context context = new PoiContext(vars);
        Object poiObject = context.getVar(PoiContext.POI_OBJECT_KEY);
        assertNotNull(poiObject);
        assertTrue(poiObject instanceof PoiUtil);
        assertEquals(123, context.getVar("a"));
    }
}
