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
    private Size size;
    private Command command;

    public CommandData(AreaRef areaRef, Command command) {
        startCellRef = areaRef.getFirstCellRef();
        size = areaRef.getSize();
        this.command = command;
        sourceStartCellRef = startCellRef;
        sourceSize = size;
    }

    public CommandData(String areaRef, Command command) {
        this(new AreaRef(areaRef), command);
    }

    public CommandData(CellRef startCellRef, Size size, Command command) {
        this.startCellRef = startCellRef;
        this.size = size;
        this.command = command;
    }

    public AreaRef getAreaRef() {
        return new AreaRef(startCellRef, size);
    }

    public CellRef getStartCellRef() {
        return startCellRef;
    }

    public Size getSize() {
        return size;
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
        size = sourceSize;
        command.reset();
    }

    void resetStartCellAndSize() {
        startCellRef = sourceStartCellRef;
        size = sourceSize;
    }
}
