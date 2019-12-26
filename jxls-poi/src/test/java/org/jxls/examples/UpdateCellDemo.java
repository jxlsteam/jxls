package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.List;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.Test;
import org.jxls.command.CellDataUpdater;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.PoiCellData;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Demonstrates jx:updateCell command
 */
public class UpdateCellDemo {
    private static final Logger logger = LoggerFactory.getLogger(UpdateCellDemo.class);

    @Test
    public void test() throws ParseException, IOException {
        // TODO MW: Demo Refactoring
        logger.info("Running UpdateCell command demo");
        List<Employee> employees = Employee.generateSampleEmployeeData();
        try (InputStream is = UpdateCellDemo.class.getResourceAsStream("updatecell_template.xlsx")) {
            try (OutputStream os = new FileOutputStream("target/updatecell_output.xlsx")) {
                Context context = new Context();
                context.putVar("employees", employees);
                context.putVar("myCellUpdater", new MyCellUpdater());
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }

    public static class MyCellUpdater implements CellDataUpdater {
        
        @Override
        public void updateCellData(CellData cellData, CellRef targetCell, Context context) {
            PoiCellData poiCellData = (PoiCellData) cellData;
            poiCellData.getCellStyle().setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        }
    }
}
