package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.AreaListener;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.PoiTransformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 1st Highlight Demo
 * 
 * @author Leonid Vysochyn Date: 10/22/13
 */
public class Highlight1Demo {
    private static final Logger logger = LoggerFactory.getLogger(Highlight1Demo.class);

    @Test
    public void test() throws ParseException, IOException {
        logger.info("Running Highlight demo");
        List<Employee> employees = Employee.generateSampleEmployeeData();
        try (InputStream is = Highlight1Demo.class.getResourceAsStream("highlight_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/highlight_output.xls")) {
                PoiTransformer transformer = PoiTransformer.createTransformer(is, os);
                AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
                List<Area> xlsAreaList = areaBuilder.build(transformer, false);
                Area mainArea = xlsAreaList.get(0);
                Area loopArea = xlsAreaList.get(0).getCommandDataList().get(0).getCommand().getAreaList().get(0);
                loopArea.addAreaListener(new HighlightCellAreaListener(transformer));
                Context context = new Context();
                context.putVar("employees", employees);
                mainArea.applyAt(new CellRef("Result!A1"), context);
                mainArea.processFormulas();
                transformer.write();
            }
        }
    }

    public static class HighlightCellAreaListener implements AreaListener {
        private static final Logger logger = LoggerFactory.getLogger(HighlightCellAreaListener.class);
        private PoiTransformer transformer;
        private final CellRef paymentCell = new CellRef("Template!C4");

        public HighlightCellAreaListener(PoiTransformer transformer) {
            this.transformer = transformer;
        }

        @Override
        public void beforeApplyAtCell(CellRef cellRef, Context context) {
        }

        @Override
        public void afterApplyAtCell(CellRef cellRef, Context context) {
        }

        @Override
        public void beforeTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
        }

        @Override
        public void afterTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
            logger.info("Source: " + srcCell.getCellName() + ", Target: " + targetCell.getCellName());
            if (paymentCell.equals(srcCell)) { // we are at employee payment cell
                Employee employee = (Employee) context.getVar("employee");
                if (employee.getPayment().doubleValue() > 2000) { // highlight payment when >= $2000
                    logger.info("highlighting payment for employee " + employee.getName());
                    highlightCell(targetCell);
                }
            }
        }

        private void highlightCell(CellRef cellRef) {
            Workbook workbook = transformer.getWorkbook();
            Sheet sheet = workbook.getSheet(cellRef.getSheetName());
            Cell cell = sheet.getRow(cellRef.getRow()).getCell(cellRef.getCol());
            CellStyle cellStyle = cell.getCellStyle();
            CellStyle newCellStyle = workbook.createCellStyle();
            newCellStyle.setDataFormat(cellStyle.getDataFormat());
            newCellStyle.setFont(workbook.getFontAt(cellStyle.getFontIndex()));
            newCellStyle.setFillBackgroundColor(cellStyle.getFillBackgroundColor());
            newCellStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
            newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(newCellStyle);
        }
    }
}
