package org.jxls.transform.poi;

import java.util.Map;

import org.jxls.common.Context;
import org.jxls.transform.SafeSheetNameBuilder;

/**
 * Context wrapper for POI
 * Automatically adds classes {@link PoiUtil} and {@link PoiSafeSheetNameBuilder} to the context.
 */
public class PoiContext extends Context {
    public static final String POI_OBJECT_KEY = "util";

    public PoiContext() {
        varMap.put(POI_OBJECT_KEY, new PoiUtil());
        varMap.put(SafeSheetNameBuilder.CONTEXT_VAR_NAME, new PoiSafeSheetNameBuilder());
    }

    public PoiContext(Map<String, Object> map) {
        super(map);
        varMap.put(POI_OBJECT_KEY, new PoiUtil());
        varMap.put(SafeSheetNameBuilder.CONTEXT_VAR_NAME, new PoiSafeSheetNameBuilder());
    }
}
