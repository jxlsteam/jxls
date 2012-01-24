package com.jxls.writer;

import com.jxls.writer.command.Context;


/**
 * Date: Apr 3, 2009
 *
 * @author Leonid Vysochyn
 */
public interface XlsProxy {

    void processCell(Cell cell, Cell newCell, Context context);

}
