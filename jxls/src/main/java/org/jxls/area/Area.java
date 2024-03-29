package org.jxls.area;

import java.util.List;

import org.jxls.command.Command;
import org.jxls.common.AreaListener;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.formula.FormulaProcessor;
import org.jxls.transform.Transformer;

/**
 * Generic interface for excel area processing
 * 
 * @author Leonid Vysochyn
 */
public interface Area {

    Size applyAt(CellRef cellRef, Context context);

    CellRef getStartCellRef();
    
    Size getSize();
    
    AreaRef getAreaRef();

    List<CommandData> getCommandDataList();

    void addCommand(AreaRef ref, Command command);

    Transformer getTransformer();
    
    void processFormulas(FormulaProcessor formulaProcessor);
    
    void addAreaListener(AreaListener listener);
    
    List<AreaListener> getAreaListeners();

    List<Command> findCommandByName(String name);

    void reset();

    Command getParentCommand();
    
    void setParentCommand(Command command);
    
    void clearCells();
}
