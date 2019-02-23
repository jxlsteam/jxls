package org.jxls.common;

import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import java.util.HashMap;
import java.util.Map;
import org.junit.Assert;
import org.junit.Test;
import org.jxls.transform.poi.PoiContext;
import org.jxls.transform.poi.PoiUtil;

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
        Map<String, Object> vars = new HashMap<String, Object>();
        vars.put("a", 123);
        Context context = new PoiContext(vars);
        Object poiObject = context.getVar(PoiContext.POI_OBJECT_KEY);
        assertNotNull(poiObject);
        assertTrue(poiObject instanceof PoiUtil);
        Assert.assertEquals(123, context.getVar("a"));
    }
}
