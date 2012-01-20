package com.jxls.writer.command;

import com.jxls.writer.*;

import java.util.List;

/**
 * @author Leonid Vysochyn
 *         Date: 21.03.2009
 */
public abstract class AbstractCommand implements Command {
    Pos pos;
    Size initialSize;

    protected AbstractCommand(Pos pos, Size initialSize) {
        this.pos = pos;
        this.initialSize = initialSize;
    }

    @Override
    public Pos getPos() {
        return pos;
    }

    @Override
    public Size getInitialSize() {
        return initialSize;
    }
}
