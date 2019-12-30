package org.jxls.examples;

import java.io.IOException;
import java.text.ParseException;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.command.CellDataUpdater;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.PoiCellData;

/**
 * Demonstrates jx:updateCell command
 */
public class UpdateCellDemo {

    @Test
    public void test() throws ParseException, IOException {
        Context context = new Context();
        context.putVar("myCellUpdater", new MyCellUpdater());
        context.putVar("employees", Employee.generateSampleEmployeeData());
        
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
    }

    public static class MyCellUpdater implements CellDataUpdater {
        
        @Override
        public void updateCellData(CellData cellData, CellRef targetCell, Context context) {
            PoiCellData poiCellData = (PoiCellData) cellData;
            poiCellData.getCellStyle().setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        }
    }
}
