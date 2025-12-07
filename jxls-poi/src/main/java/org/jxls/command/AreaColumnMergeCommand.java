package org.jxls.command;

import org.jxls.area.Area;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;

public class AreaColumnMergeCommand extends AbstractMergeCellsCommand {
    public static final String COMMAND_NAME = "areaColumnMerge";

    @Override
    public String getName() {
        return COMMAND_NAME;
    }

    @Override
    public Command addArea(Area area) {
        if (area == null) {
            return this;
        }
        if (area.getStartCellRef().getRow() != area.getAreaRef().getLastCellRef().getRow()) {
            throw new IllegalArgumentException("You can add only a single row to '" + COMMAND_NAME + "' command");
        }
        return super.addArea(area);
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        Area area = getArea();
        area.applyAt(cellRef, context);
        int rowsToMerge = evaluate(getRows(), cellRef, context);
        if (rowsToMerge > 1) {
            int startRow = cellRef.getRow();
            int startCol = cellRef.getCol();
            Size size = area.getAreaRef().getSize();
            int width = size.getWidth();
            for (int i = 0; i < width; i++) {
                CellRef ref = new CellRef(cellRef.getSheetName(), startRow, startCol + i);
                mergeCells(ref, rowsToMerge, 1);
            }
        }
        return new Size(area.getSize().getWidth(), area.getSize().getHeight() + rowsToMerge - 1);
    }
}
