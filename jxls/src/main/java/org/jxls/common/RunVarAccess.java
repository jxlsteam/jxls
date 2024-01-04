package org.jxls.common;

import java.util.Map;

public interface RunVarAccess {

    Object getRunVar(String name, Map<String, Object> data);
}
