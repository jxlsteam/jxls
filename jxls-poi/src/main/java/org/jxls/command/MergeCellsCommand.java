package org.jxls.command;

import org.jxls.area.Area;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;

/**
 * <p>Merge cells</p>
 * <pre>jx:mergeCells(
 * lastCell="Merge cell ranges"
 * [, cols="Number of columns combined"]
 * [, rows="Number of rows combined"]
 * [, minCols="Minimum number of columns to merge"]
 * [, minRows="Minimum number of rows to merge"]
 * )</pre>
 * <p>Note: this command can only be used on cells that have not been merged. An exception will occur if the scope of
 * the merged cell exists for the merged cell</p>
 *
 * @author lnk
 * @since 2.6.0
 */
public class MergeCellsCommand extends AbstractMergeCellsCommand {
    public static final String COMMAND_NAME = "mergeCells";

    /** Number of columns combined */
    private String cols;
    /** Minimum number of columns to merge */
    private String minCols;
    /** Minimum number of rows to merge */
    private String minRows;

    @Override
    public String getName() {
        return COMMAND_NAME;
    }

    public String getCols() {
        return cols;
    }

    public void setCols(String cols) {
        this.cols = cols;
    }

    public String getMinCols() {
        return minCols;
    }

    public void setMinCols(String minCols) {
        this.minCols = minCols;
    }

    public String getMinRows() {
        return minRows;
    }

    public void setMinRows(String minRows) {
        this.minRows = minRows;
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        Area area = getArea();
        int rows = evaluate(getRows(), cellRef, context);
        int cols = evaluate(this.cols, cellRef, context);
        rows = Math.max(evaluate(this.minRows, cellRef, context), rows);
        cols = Math.max(evaluate(this.minCols, cellRef, context), cols);
        rows = rows > 0 ? rows : area.getSize().getHeight();
        cols = cols > 0 ? cols : area.getSize().getWidth();
        if (rows > 1 || cols > 1) {
            mergeCells(cellRef, rows, cols);
        }
        area.applyAt(cellRef, context);
        return new Size(cols, rows);
    }
}
