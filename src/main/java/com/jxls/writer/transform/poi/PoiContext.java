package com.jxls.writer.transform.poi;

import com.jxls.writer.common.Context;

import java.util.Map;

/**
 * Context wrapper for POI
 * Automatically adds {@link PoiUtil} class to the context
 * @author Leonid Vysochyn
 *         Date: 6/18/12
 */
public class PoiContext extends Context {

    public static final String POI_OBJECT_KEY = "poi";

    public PoiContext() {
        varMap.put(POI_OBJECT_KEY, new PoiUtil());
    }

    public PoiContext(Map<String, Object> map) {
        super(map);
        varMap.put(POI_OBJECT_KEY, new PoiUtil());
    }
}
