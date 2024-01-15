package org.jxls.transform;

import org.jxls.common.PublicContext;

public interface PreWriteAction {

    void preWrite(Transformer transformer, PublicContext context);
}
