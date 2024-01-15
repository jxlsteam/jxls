package org.jxls.examples;

import java.io.IOException;
import java.text.ParseException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.area.Area;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.AreaListener;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;
import org.jxls.transform.poi.PoiTransformer;

/**
 * @author Leonid Vysochyn 
 * @since 10/22/13
 */
public class HighlightDemo {

    @Test
    public void test() throws ParseException, IOException {
        // Prepare
        Map<String,Object> data = new HashMap<>();
        data.put("employees", Employee.generateSampleEmployeeData());
        
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance().withAreaBuilder(new XlsCommentAreaBuilder() {
            @Override
            public List<Area> build(Transformer transformer, boolean ctc) {
                List<Area> areas = super.build(transformer, ctc);
                addAreaListener(new HighlightCellAreaListener((PoiTransformer) transformer),
                        new AreaRef("Template!A4:D4")/*jx:each command*/, areas);
                return areas;
            }
        }));
        
        // Verify: cells C5, C6 and C8 must have a orange background.
    }

    public static class HighlightCellAreaListener implements AreaListener {
        private final PoiTransformer transformer;
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
            if (paymentCell.equals(srcCell)) { // we are at employee payment cell
                Employee employee = (Employee) context.getVar("employee");
                if (employee.getPayment().doubleValue() > 2000) { // highlight payment when >= $2000
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
