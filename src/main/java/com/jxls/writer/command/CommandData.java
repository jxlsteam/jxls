package com.jxls.writer.command;

import com.jxls.writer.Pos;

/**
 * @author Leonid Vysochyn
 *         Date: 1/25/12 1:12 PM
 */
public class CommandData {
    Pos startPos;
    Command command;

    public CommandData(Pos startPos, Command command) {
        this.startPos = startPos;
        this.command = command;
    }

    public Pos getStartPos() {
        return startPos;
    }

    public Command getCommand() {
        return command;
    }

    public void setStartPos(Pos startPos) {
        this.startPos = startPos;
    }
}
