package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.area.XlsArea;
import org.jxls.command.Command;
import org.jxls.command.EachCommand;
import org.jxls.command.IfCommand;
import org.jxls.common.AreaListener;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.entity.Department;
import org.jxls.entity.Employee;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;

/**
 * @author Leonid Vysochyn
 * @since 2/16/12
 */
public class AreaListenerDemo {
    private static final String template = "each_if_demo.xls";
    private static final String output = "target/listener_demo_output.xls";

    @Test
    public void test() throws IOException {
        List<Department> departments = Department.createDepartments();
        try (InputStream is = EachIfCommandDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = Jxls3Tester.createTransformer(is, os); 
                XlsArea xlsArea = new XlsArea("Template!A1:G15", transformer);
                XlsArea departmentArea = new XlsArea("Template!A2:G12", transformer);
                EachCommand departmentEachCommand = new EachCommand("department", "departments", departmentArea);
                XlsArea employeeArea = new XlsArea("Template!A9:F9", transformer);
                XlsArea ifArea = new XlsArea("Template!A18:F18", transformer);
                XlsArea elseArea = new XlsArea("Template!A9:F9", transformer);
                IfCommand ifCommand = new IfCommand("employee.payment <= 2000", ifArea, elseArea);
                ifArea.addAreaListener(new SimpleAreaListener(ifArea));
                elseArea.addAreaListener(new SimpleAreaListener(elseArea));
                employeeArea.addCommand(new AreaRef("Template!A9:F9"), ifCommand);
                Command employeeEachCommand = new EachCommand("employee", "department.staff", employeeArea);
                departmentArea.addCommand(new AreaRef("Template!A9:F9"), employeeEachCommand);
                xlsArea.addCommand(new AreaRef("Template!A2:F12"), departmentEachCommand);
                Context context = new Context();
                context.putVar("departments", departments);
                xlsArea.applyAt(new CellRef("Down!A1"), context);
                xlsArea.setFormulaProcessor(new StandardFormulaProcessor());
                xlsArea.processFormulas();
                departmentEachCommand.setDirection(EachCommand.Direction.RIGHT);
                xlsArea.reset();
                xlsArea.applyAt(new CellRef("Right!A1"), context);
                xlsArea.processFormulas();
                transformer.write();
            }
        }
    }

    public static class SimpleAreaListener implements AreaListener {
        private final CellRef bonusCell1 = new CellRef("Template!E9");
        private final CellRef bonusCell2 = new CellRef("Template!E18");
        private PoiTransformer transformer;

        public SimpleAreaListener(XlsArea area) {
            transformer = (PoiTransformer) area.getTransformer();
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
            if (bonusCell1.equals(srcCell) || bonusCell2.equals(srcCell)) { // we are at employee bonus cell
                Employee employee = (Employee) context.getVar("employee");
                if (employee.getBonus().doubleValue() >= 0.2) { // highlight bonus when >= 20%
                    highlightBonus(targetCell);
                }
            }
        }

        private void highlightBonus(CellRef cellRef) {
            Workbook workbook = transformer.getWorkbook();
            Sheet sheet = workbook.getSheet(cellRef.getSheetName());
            Cell cell = sheet.getRow(cellRef.getRow()).getCell(cellRef.getCol());
            CellStyle cellStyle = cell.getCellStyle();
            CellStyle newCellStyle = workbook.createCellStyle();
            newCellStyle.setDataFormat(cellStyle.getDataFormat());
            newCellStyle.setFont(workbook.getFontAt(cellStyle.getFontIndex()));
            newCellStyle.setFillBackgroundColor(cellStyle.getFillBackgroundColor());
            newCellStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
            // newCellStyle.setFillForegroundColor( cellStyle.getFillForegroundColor());
            newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(newCellStyle);
        }
    }
}
