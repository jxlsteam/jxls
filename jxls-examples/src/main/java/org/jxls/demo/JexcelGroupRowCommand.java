package org.jxls.demo;

import org.jxls.area.Area;
import org.jxls.command.AbstractCommand;
import org.jxls.command.Command;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.transform.jexcel.JexcelTransformer;
import org.jxls.util.Util;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn
 *         Date: 11/15/13
 */
public class JexcelGroupRowCommand extends AbstractCommand {
    static Logger logger = LoggerFactory.getLogger(JexcelGroupRowCommand.class);
    Area area;
    String collapseIf;

    public String getName() {
        return "groupRow";
    }

    public Size applyAt(CellRef cellRef, Context context) {
        Size resultSize = area.applyAt(cellRef, context);
        if( resultSize.equals(Size.ZERO_SIZE)) return resultSize;
        int startRow = cellRef.getRow();
        int endRow = cellRef.getRow() + resultSize.getHeight() - 1;
        try{
            JexcelTransformer transformer = (JexcelTransformer) area.getTransformer();
            WritableWorkbook workbook = transformer.getWritableWorkbook();
            WritableSheet sheet = workbook.getSheet(cellRef.getSheetName());
            boolean collapseFlag = false;
            if( collapseIf != null && collapseIf.trim().length() > 0){
                collapseFlag = Util.isConditionTrue(getTransformationConfig().getExpressionEvaluator(), collapseIf, context);
            }
            sheet.setRowGroup(startRow, endRow, collapseFlag);
        }catch(Exception e){
            logger.error("Failed to apply JexcelGroupRowCommand at " + cellRef, e);
        }
        return resultSize;
    }

    @Override
    public Command addArea(Area area) {
        super.addArea(area);
        this.area = area;
        return this;
    }

    public void setCollapseIf(String collapseIf) {
        this.collapseIf = collapseIf;
    }
}
