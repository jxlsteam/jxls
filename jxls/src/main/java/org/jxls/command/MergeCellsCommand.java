package org.jxls.command;

import org.jxls.area.Area;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.transform.Transformer;

/**
 * <p>Merge cells</p>
 * jx:merge(
 * lastCell="Merge cell ranges"
 * [, cols="Number of columns combined"]
 * [, rows="Number of rows combined"]
 * [, minCols="Minimum number of columns to merge"]
 * [, minRows="Minimum number of rows to merge"]
 * )
 * Note: this command can only be used on cells that have not been merged. An exception will occur if the scope of
 * the merged cell exists for the merged cell
 *
 * @Author lnk
 * @Date 2019/2/14
 */
public class MergeCellsCommand extends AbstractCommand {
    public static final String COMMAND_NAME = "mergeCells";

    private String cols;        //Number of columns combined
    private String rows;        //Number of rows combined
    private String minCols;     //Minimum number of columns to merge
    private String minRows;     //Minimum number of rows to merge

    private Area area;

    @Override
    public String getName() {
        return COMMAND_NAME;
    }

    @Override
    public Command addArea(Area area) {
        if (super.getAreaList().size() >= 1) {
            throw new IllegalArgumentException("You can add only a single area to 'merge' command");
        }
        this.area = area;
        return super.addArea(area);
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        int rows = getVal(this.rows, context);
        int cols = getVal(this.cols, context);
        rows = Math.max(getVal(this.minRows, context), rows);
        cols = Math.max(getVal(this.minCols, context), cols);
        rows = rows > 0 ? rows : area.getSize().getHeight();
        cols = cols > 0 ? cols : area.getSize().getWidth();
        if (rows > 1 || cols > 1) {
            Transformer transformer = this.getTransformer();
            transformer.mergeCells(cellRef, rows, cols);
        }
        area.applyAt(cellRef, context);
        return new Size(cols, rows);
    }

    private int getVal(String expression, Context context) {
        if (expression != null && expression.trim().length() > 0) {
            Object obj = getTransformationConfig().getExpressionEvaluator().evaluate(expression, context.toMap());
            try {
                return Integer.parseInt(obj.toString());
            } catch (NumberFormatException e) {
                throw new IllegalArgumentException("Expression: " + expression + " failed to resolve");
            }
        }
        return 0;
    }

    public String getCols() {
        return cols;
    }

    public void setCols(String cols) {
        this.cols = cols;
    }

    public String getRows() {
        return rows;
    }

    public void setRows(String rows) {
        this.rows = rows;
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
}
