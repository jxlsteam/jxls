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
    private int[] blankLinesUp;
    private int tmpMinBlankLines; // ok, I know it's bad here. Temporary result of calcMinBlankLines

    public CommandData(AreaRef areaRef, Command command) {
        startCellRef = areaRef.getFirstCellRef();
        size = areaRef.getSize();
        this.command = command;
        sourceStartCellRef = startCellRef;
        sourceSize = size;
        blankLinesUp = new int[startCellRef.getCol() + size.getWidth()];
    }

    public CommandData(String areaRef, Command command) {
        this(new AreaRef(areaRef), command);
    }

    public CommandData(CellRef startCellRef, Size size, Command command) {
        this.startCellRef = startCellRef;
        this.size = size;
        this.command = command;
        blankLinesUp = new int[startCellRef.getCol() + size.getWidth()];
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

    public int[] getBlankLinesUp() {
        return blankLinesUp;
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
        // to do or not to do ?!? blankCellsUp = new int[size.getWidth()];
    }

    public int getTmpMinBlankLines() {
        return tmpMinBlankLines;
    }

    public void setTmpMinBlankLines(int tmpMinBlankLines) {
        this.tmpMinBlankLines = tmpMinBlankLines;
    }

    /**
     * Calculate minimum numbare of blank lines up to columns between 
     * startCol and endCol (inclusive)
     * All arguments are 0-based and relative to containing area
     *
     * @param startCol 
     * @param endCol
     * @return
     */
    public int calcMinBlankLines(int startCol, int endCol) {
        int minBlankLines = Integer.MAX_VALUE;
        for (int col = startCol; col <= endCol; col++) {
            minBlankLines = Math.min(minBlankLines, blankLinesUp[col]);
        }
        return minBlankLines == Integer.MAX_VALUE ? 0 : minBlankLines;
    }

    /**
     * Add blank lines to all columns *outside* of [startCol, endCol] range
     * All arguments are 0-based and relative to containing area
     */
    public void addBlankLines(int realHeightChange, int startCol, int endCol) {
        for (int col = 0; col < blankLinesUp.length; col++) {
            if (col < startCol || col > endCol) {
                blankLinesUp[col] += realHeightChange;
            }
        }
    }
}
