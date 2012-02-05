package com.jxls.writer.command;

import com.jxls.writer.Pos;
import com.jxls.writer.transform.Transformer;

/**
 * @author Leonid Vysochyn
 *         Date: 1/18/12 5:20 PM
 */
public interface Area extends Command{
    Pos getStartPos();
    Transformer getTransformer();
    void processFormulas();
}
