package com.jxls.writer.command;

import com.jxls.writer.AreaRef;
import com.jxls.writer.CellRef;
import com.jxls.writer.Size;

/**
 * @author Leonid Vysochyn
 *         Date: 1/25/12 1:12 PM
 */
public class CommandData {
    CellRef startCellRef;
    Size size;
    Command command;
    
    public CommandData(AreaRef areaRef, Command command){
        startCellRef = areaRef.getFirstCellRef();
        size = areaRef.getSize();
        this.command = command;
    }
    
    public CommandData(String areaRef, Command command){
        this(new AreaRef(areaRef), command);
    }

    public CommandData(CellRef startCellRef, Size size, Command command) {
        this.startCellRef = startCellRef;
        this.size = size;
        this.command = command;
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
}
