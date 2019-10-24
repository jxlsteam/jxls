package org.jxls.demo;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.jxls.command.CellDataUpdater;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.demo.guide.Employee;
import org.jxls.demo.guide.ObjectCollectionDemo;
import org.jxls.transform.poi.PoiCellData;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.List;

/**
 * Demonstrates jx:updateCell command
 */
public class UpdateCellDemo {
    private static Logger logger = LoggerFactory.getLogger(UpdateCellDemo.class);

    public static void main(String[] args) throws ParseException, IOException {
        logger.info("Running UpdateCell command demo");

        List<Employee> employees = ObjectCollectionDemo.generateSampleEmployeeData();
        try(InputStream is = UpdateCellDemo.class.getResourceAsStream("updatecell_template.xlsx")) {
            try (OutputStream os = new FileOutputStream("target/updatecell_output.xlsx")) {
                Context context = new Context();
                context.putVar("employees", employees);
                context.putVar("myCellUpdater", new MyCellUpdater());
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }

    static class MyCellUpdater implements CellDataUpdater {
        @Override
        public void updateCellData(CellData cellData, CellRef targetCell, Context context) {
            PoiCellData poiCellData = (PoiCellData) cellData;
            poiCellData.getCellStyle().setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        }
    }
}
