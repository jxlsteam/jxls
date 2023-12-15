package org.jxls.command;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jxls.area.Area;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.transform.poi.PoiTransformer;

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
public class MergeCellsCommand extends AbstractCommand {
    public static final String COMMAND_NAME = "mergeCells";

    /** Number of columns combined */
    private String cols;
    /** Number of rows combined */
    private String rows;
    /** Minimum number of columns to merge */
    private String minCols;
    /** Minimum number of rows to merge */
    private String minRows;
    private Area area;

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

    @Override
    public Command addArea(Area area) {
        if (super.getAreaList().size() >= 1) {
            throw new IllegalArgumentException("You can only add 1 area to 'mergeCells' command!");
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
            mergeCells(((PoiTransformer) getTransformer()).getWorkbook(), cellRef, rows, cols);
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
    
    private void mergeCells(Workbook workbook, CellRef cellRef, int rows, int cols) {
        Sheet sheet = workbook.getSheet(cellRef.getSheetName());
        CellRangeAddress region = new CellRangeAddress(
                cellRef.getRow(),
                cellRef.getRow() + rows - 1,
                cellRef.getCol(),
                cellRef.getCol() + cols - 1);
        sheet.addMergedRegion(region);

        CellStyle cellStyle = null;
		try {
            cellStyle = ((PoiTransformer) getTransformer()).getCellStyle(cellRef);
        } catch (Exception ignore) {
        }
        for (int i = region.getFirstRow(); i <= region.getLastRow(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                row = sheet.createRow(i);
            }
            for (int j = region.getFirstColumn(); j <= region.getLastColumn(); j++) {
                Cell cell = row.getCell(j);
                if (cell == null) {
                    cell = row.createCell(j);
                }
                if (cellStyle == null) {
                    cell.getCellStyle().setAlignment(HorizontalAlignment.CENTER);
                    cell.getCellStyle().setVerticalAlignment(VerticalAlignment.CENTER);
                } else {
                    cell.setCellStyle(cellStyle);
                }
            }
        }
    }
}
