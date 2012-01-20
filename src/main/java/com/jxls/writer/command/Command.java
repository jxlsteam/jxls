package com.jxls.writer.command;

import com.jxls.writer.*;

/**
 * Date: Mar 13, 2009
 *
 * @author Leonid Vysochyn
 */
public interface Command {

    Pos getPos();

    Size getInitialSize();

    Size applyAt(Pos pos, Context context);

    Size getSize(Context context);

}
