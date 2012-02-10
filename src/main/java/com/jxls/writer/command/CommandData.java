package com.jxls.writer.command;

import com.jxls.writer.Pos;
import com.jxls.writer.Size;

/**
 * @author Leonid Vysochyn
 *         Date: 1/25/12 1:12 PM
 */
public class CommandData {
    Pos startPos;
    Size size;
    Command command;

    public CommandData(Pos startPos, Size size, Command command) {
        this.startPos = startPos;
        this.size = size;
        this.command = command;
    }

    public CommandData(Pos startPos, Command command) {
        this.startPos = startPos;
        this.command = command;
    }

    public Pos getStartPos() {
        return startPos;
    }

    public Size getSize() {
        return size;
    }

    public Command getCommand() {
        return command;
    }

    public void setStartPos(Pos startPos) {
        this.startPos = startPos;
    }
}
