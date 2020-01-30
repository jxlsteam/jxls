package org.jxls.unittests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import java.util.HashMap;
import java.util.Map;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.transform.poi.PoiContext;
import org.jxls.transform.poi.PoiUtil;

/**
 * @author Leonid Vysochyn
 * @since 6/18/12 5:13 PM
 */
public class PoiContextTest {
    
    @Test
    public void createEmptyContext() {
        Context context = new PoiContext();
        Object poiObject = context.getVar(PoiContext.POI_OBJECT_KEY);
        assertNotNull(poiObject);
        assertTrue(poiObject instanceof PoiUtil);
    }

    @Test
    public void createNonEmptyContext() {
        Map<String, Object> vars = new HashMap<>();
        vars.put("a", 123);
        Context context = new PoiContext(vars);
        Object poiObject = context.getVar(PoiContext.POI_OBJECT_KEY);
        assertNotNull(poiObject);
        assertTrue(poiObject instanceof PoiUtil);
        assertEquals(123, context.getVar("a"));
    }
}
