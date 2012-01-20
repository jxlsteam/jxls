package com.jxls.writer;

import com.jxls.writer.Pos;
import com.jxls.writer.command.Context;


/**
 * Date: Apr 3, 2009
 *
 * @author Leonid Vysochyn
 */
public interface XlsProxy {

    void processCell(Pos pos, Pos newPos, Context context);

}
