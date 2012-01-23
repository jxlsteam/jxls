package com.jxls.writer.transform;

import com.jxls.writer.Pos;
import com.jxls.writer.command.Context;

/**
 * @author Leonid Vysochyn
 *         Date: 1/23/12 1:24 PM
 */
public interface Transformer {
    void transform(Pos pos, Pos newPos, Context context);
}
