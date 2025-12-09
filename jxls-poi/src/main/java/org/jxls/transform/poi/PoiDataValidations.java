package org.jxls.transform.poi;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.jxls.area.Area;
import org.jxls.command.EachCommand.Direction;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;
import org.jxls.common.Size;
import org.jxls.logging.JxlsLogger;

public class PoiDataValidations {
    public static boolean FEATURE_TOGGLE = true;

    public void dataValidation(Area area, Size size, Direction direction, Sheet sheet, JxlsLogger logger) {
        AreaRef areaRef = area.getAreaRef();
        CellRef first = areaRef.getFirstCellRef();
        CellRef last = areaRef.getLastCellRef();
        for (DataValidation v : sheet.getDataValidations()) {
            for (CellRangeAddress region : v.getRegions().getCellRangeAddresses()) {
                int vRow = region.getFirstRow();
                int vCol = region.getFirstColumn();
                if (region.getNumberOfCells() == 1 && areaRef.contains(vRow, vCol)) {
                    // DataValidation found
                    switch (direction) {
                    case DOWN:
                        down(size, sheet, v, region, first, last, vRow, vCol);
                        break;
                    case RIGHT:
                        right(size, sheet, v, region, first, last, vRow, vCol);
                        break;
                    }
                }
            }
        }
    }
    
    private void down(Size size, Sheet sheet, DataValidation v, CellRangeAddress region,
            CellRef first, CellRef last, int row, int col) {
        int height = last.getRow() - first.getRow() + 1;
        if (height == 1) {
            // combined area
            copyDataValidation(v, sheet, new CellRangeAddressList(row + 1, row + size.getHeight() - 1, col, col));
        } else {
            for (int r = row + height; r < first.getRow() + size.getHeight(); r += height) {
                copyDataValidation(v, sheet, new CellRangeAddressList(r, r, col, col));
            }
        }
    }

    private void right(Size size, Sheet sheet, DataValidation v, CellRangeAddress region,
            CellRef first, CellRef last, int row, int col) {
        int width = last.getCol() - first.getCol() + 1;
        if (width == 1) {
            // combined area
            copyDataValidation(v, sheet, new CellRangeAddressList(row, row, col + 1, col + size.getWidth()));
        } else {
            for (int c = col + width; c < first.getCol() + size.getWidth(); c += width) {
                copyDataValidation(v, sheet, new CellRangeAddressList(row, row, c, c));
            }
        }
    }

    private void copyDataValidation(DataValidation v, Sheet sheet, CellRangeAddressList addressList) {
        DataValidation nv = sheet.getDataValidationHelper().createValidation(v.getValidationConstraint(), addressList);
        nv.setErrorStyle(v.getErrorStyle());
        nv.setShowErrorBox(v.getShowErrorBox());
        nv.createErrorBox(v.getErrorBoxTitle(), v.getErrorBoxText());
        nv.setEmptyCellAllowed(v.getEmptyCellAllowed());
        nv.setSuppressDropDownArrow(v.getSuppressDropDownArrow());
        sheet.addValidationData(nv);
    }
}
