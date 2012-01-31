package com.jxls.writer.command;

import com.jxls.writer.*;

/**
 * @author Leonid Vysochyn
 *         Date: 21.03.2009
 */
public abstract class AbstractCommand implements Command {
    Size initialSize;

    protected AbstractCommand(Size initialSize) {
        this.initialSize = initialSize;
    }

    public Size getInitialSize() {
        return initialSize;
    }
}
