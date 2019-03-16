package org.jxls.command;

import org.jxls.area.Area;
import org.jxls.common.Size;
import org.jxls.common.CellRef;
import org.jxls.common.Context;

import java.util.List;

/**
 * A command interface defines a transformation of a list of areas at a specified cell
 *
 * @author Leonid Vysochyn
 * @since Mar 13, 2009
 */
public interface Command {
    String INNER_SHIFT_MODE = "inner";
    String ADJACENT_SHIFT_MODE = "adjacent";

    /**
     * @return command name
     */
    String getName();

    /**
     * @return a list of areas for this command
     */
    List<Area> getAreaList();

    /**
     * Adds an area to this command
     * 
     * @param area to be added area
     * @return this command instance
     */
    Command addArea(Area area);

    /**
     * Applies a command at the given cell reference
     * 
     * @param cellRef cell reference where the command must be applied
     * @param context bean context to use
     * @return size of enclosing command area after transformation
     */
    Size applyAt(CellRef cellRef, Context context);

    /**
     * Resets command data for repeatable command usage
     */
    void reset();

    // TODO MW to Leonid: Javadoc
    void setShiftMode(String mode);

    // TODO MW to Leonid: Javadoc
    String getShiftMode();
}
