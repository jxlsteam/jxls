package org.jxls.area;

import org.jxls.common.Size;
import org.jxls.command.Command;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;

/**
 * A command holder class
 * 
 * @author Leonid Vysochyn
 */
public class CommandData {
    private CellRef sourceStartCellRef;
    private Size sourceSize;
    private CellRef startCellRef;
    private Size commandSize;
    private Command command;

    public CommandData(AreaRef areaRef, Command command) {
        startCellRef = areaRef.getFirstCellRef();
        commandSize = areaRef.getSize();
        this.command = command;
        sourceStartCellRef = startCellRef;
        sourceSize = commandSize;
    }

    public CommandData(String areaRef, Command command) {
        this(new AreaRef(areaRef), command);
    }

    public CommandData(CellRef startCellRef, Size size, Command command) {
        this.startCellRef = startCellRef;
        this.commandSize = size;
        this.command = command;
    }

    public AreaRef getAreaRef() {
        return new AreaRef(startCellRef, commandSize);
    }

    public CellRef getStartCellRef() {
        return startCellRef;
    }

    public Size getCommandSize() {
        return commandSize;
    }

    public Command getCommand() {
        return command;
    }

    public void setStartCellRef(CellRef startCellRef) {
        this.startCellRef = startCellRef;
    }

    public CellRef getSourceStartCellRef() {
        return sourceStartCellRef;
    }

    public void setSourceStartCellRef(CellRef sourceStartCellRef) {
        this.sourceStartCellRef = sourceStartCellRef;
    }

    public Size getSourceSize() {
        return sourceSize;
    }

    public void setSourceSize(Size sourceSize) {
        this.sourceSize = sourceSize;
    }

    void reset() {
        startCellRef = sourceStartCellRef;
        commandSize = sourceSize;
        command.reset();
    }

    void resetStartCellAndSize() {
        startCellRef = sourceStartCellRef;
        commandSize = sourceSize;
    }
}
