package org.jxls.area;

import org.jxls.command.Command;
import org.jxls.formula.FormulaProcessor;
import org.jxls.transform.Transformer;
import org.jxls.common.*;

import java.util.List;

/**
 * Generic interface for excel area processing
 * @author Leonid Vysochyn
 *         Date: 1/18/12
 */
public interface Area {
    Size applyAt(CellRef cellRef, Context context);

    CellRef getStartCellRef();
    Size getSize();
    
    AreaRef getAreaRef();

    List<CommandData> getCommandDataList();

    FormulaProcessor getFormulaProcessor();

    void setFormulaProcessor(FormulaProcessor formulaProcessor);

    void addCommand(AreaRef ref, Command command);

    Transformer getTransformer();
    void processFormulas();
    
    void addAreaListener(AreaListener listener);
    List<AreaListener> getAreaListeners();

    List<Command> findCommandByName(String name);

    void reset();
}
