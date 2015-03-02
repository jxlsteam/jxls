package org.jxls.area;

import org.jxls.common.Size;
import org.jxls.command.Command;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;

/**
 * A command holder class
 * @author Leonid Vysochyn
 *         Date: 1/25/12 1:12 PM
 */
public class CommandData {
    private CellRef startCellRefCopy;
    private Size sizeCopy;

    CellRef startCellRef;
    Size size;
    Command command;
    
    public CommandData(AreaRef areaRef, Command command){
        startCellRef = areaRef.getFirstCellRef();
        size = areaRef.getSize();
        this.command = command;
        startCellRefCopy = startCellRef;
        sizeCopy = size;
    }
    
    public CommandData(String areaRef, Command command){
        this(new AreaRef(areaRef), command);
    }

    public CommandData(CellRef startCellRef, Size size, Command command) {
        this.startCellRef = startCellRef;
        this.size = size;
        this.command = command;
    }
    
    public AreaRef getAreaRef(){
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

    void reset(){
        startCellRef = startCellRefCopy;
        size = sizeCopy;
        command.reset();
    }
}
