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
import org.jxls.transform.poi.PoiTransformer;

public abstract class AbstractMergeCellsCommand extends AbstractCommand {
    private Area area;
    private String rows;
    
    public Area getArea() {
        return area;
    }

    public void setArea(Area area) {
        this.area = area;
    }
    
    public String getRows() {
        return rows;
    }

    public void setRows(String rows) {
        this.rows = rows;
    }

    @Override
    public Command addArea(Area area) {
        if (super.getAreaList().size() >= 1) {
            throw new IllegalArgumentException("You can only add 1 area to '" + getName() + "' command!");
        }
        this.area = area;
        return super.addArea(area);
    }

    protected void mergeCells(CellRef cellRef, int rows, int cols) {
        Workbook workbook = ((PoiTransformer) getTransformer()).getWorkbook();
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

    protected int evaluate(String expression, CellRef cellRef, Context context) {
        if (expression != null && !expression.isBlank()) {
            try {
                Object result = context.evaluate(expression);
                return Integer.parseInt(result.toString());
            } catch (Exception e) {
                getLogger().handleEvaluationException(e, cellRef.toString(), expression);
            }
        }
        return 0;
    }
}
