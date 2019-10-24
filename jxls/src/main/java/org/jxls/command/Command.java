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

    /**
     * Defines the {@link org.jxls.common.cellshift.CellShiftStrategy} to use
     * when shifting the cells for the command while transforming an area
     * The mode value "inner" sets the {@link org.jxls.common.cellshift.InnerCellShiftStrategy} (default)
     * The mode value "adjacent" sets the {@link org.jxls.common.cellshift.AdjacentCellShiftStrategy}
     * @param mode cell shifting mode with possible values "inner" or "adjacent"
     */
    void setShiftMode(String mode);

    /**
     * Returns the cell shifting mode defining
     * the {@link org.jxls.common.cellshift.CellShiftStrategy} to use for the command
     * Possible values are
     *      "inner" defining the {@link org.jxls.common.cellshift.InnerCellShiftStrategy} to use
     *      "adjacent" defining the {@link org.jxls.common.cellshift.AdjacentCellShiftStrategy} to use
     *      null value means the default strategy will be used ({@link org.jxls.common.cellshift.InnerCellShiftStrategy})
     * @return cell shifting mode ("inner" or "adjacent")
     */
    String getShiftMode();

    /**
     * Other commands will no longer execute in this area after locking
     * @param isLock Whether the command area is locked, value should be "true","false",null
     */
    void setLockRange(String isLock);

    /**
     * Whether the command area is locked
     * Other commands will no longer execute in this area after locking
     * @return true or false (default true)
     */
    Boolean getLockRange();
}
