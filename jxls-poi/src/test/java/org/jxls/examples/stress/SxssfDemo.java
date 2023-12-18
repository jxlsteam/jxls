package org.jxls.examples.stress;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.CellDataUpdater;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.entity.Department;
import org.jxls.entity.Employee;
import org.jxls.transform.poi.PoiContext;
import org.jxls.transform.poi.PoiTransformer;

/**
 * @author Leonid Vysochyn
 */
public class SxssfDemo {
    public static final int EMPLOYEE_COUNT = 30000;
    public static final int DEPARTMENT_COUNT = 100;
    public static final int DEP_EMPLOYEE_COUNT = 500;

    @Test
    public void test() throws IOException, InvalidFormatException, ParseException {
        simpleSxssf();
        executeStress1();
        executeStress2();
    }

    private void simpleSxssf() throws ParseException, IOException, InvalidFormatException {
        try (InputStream is = SxssfDemo.class.getResourceAsStream("sxssf_template.xlsx")) {
            List<Employee> employees = Employee.generate(10);
            try (OutputStream os = new FileOutputStream("target/simple_sxssf_output.xlsx")) {
                Workbook workbook = WorkbookFactory.create(is);
                PoiTransformer transformer = PoiTransformer.createSxssfTransformer(workbook, 5, false);
                AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
                List<Area> xlsAreaList = areaBuilder.build(transformer, true);
                Area xlsArea = xlsAreaList.get(0);
                Context context = new Context();
                context.putVar("cellRefUpdater", new CellRefUpdater());
                context.putVar("employees", employees);
                xlsArea.applyAt(new CellRef("Result!A1"), context);
                context.getConfig().setIsFormulaProcessingRequired(false); // with SXSSF you cannot use normal formula
                                                                           // processing
                workbook.setForceFormulaRecalculation(true);
                workbook.setActiveSheet(1);
                transformer.getWorkbook().write(os);
            }
        }
    }

    private void executeStress1() throws IOException, InvalidFormatException {
        List<Employee> employees = Employee.generate(EMPLOYEE_COUNT);
        try (InputStream is = SxssfDemo.class.getResourceAsStream("stress1_sxssf.xlsx")) {
            assert is != null;
            Workbook workbook = WorkbookFactory.create(is);
            PoiTransformer transformer = PoiTransformer.createSxssfTransformer(workbook);
            AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
            List<Area> xlsAreaList = areaBuilder.build(transformer, true);
            Area xlsArea = xlsAreaList.get(0);
            Context context = new PoiContext();
            context.putVar("cellRefUpdater", new CellRefUpdater());
            context.putVar("totalCellUpdater", new TotalCellUpdater());
            context.putVar("employees", employees);
            context.getConfig().setIsFormulaProcessingRequired(false);
            xlsArea.applyAt(new CellRef("Result!A1"), context);
            workbook.setForceFormulaRecalculation(true);
            workbook.setActiveSheet(1);
            try (OutputStream os = new FileOutputStream("target/sxssf_stress1_output.xlsx")) {
                transformer.getWorkbook().write(os);
            }
        }
    }

    private void executeStress2() throws IOException, InvalidFormatException {
        List<Department> departments = Department.generate(DEPARTMENT_COUNT, DEP_EMPLOYEE_COUNT);
        try (InputStream is = SxssfDemo.class.getResourceAsStream("stress2_sxssf.xlsx")) {
            Workbook workbook = WorkbookFactory.create(is);
            // setting rowAccessWindowSize to 600 to be able to process static cells in a
            // single iteration
            PoiTransformer transformer = PoiTransformer.createSxssfTransformer(workbook, 600, true);
            AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
            List<Area> xlsAreaList = areaBuilder.build(transformer, true);
            Area xlsArea = xlsAreaList.get(0);
            Context context = new PoiContext();
            context.getConfig().setIsFormulaProcessingRequired(false);
            context.putVar("cellRefUpdater", new CellRefUpdater());
            context.putVar("totalCellUpdater", new DepartmentTotalCellUpdater());
            context.putVar("departments", departments);
            xlsArea.applyAt(new CellRef("Result!A1"), context);
            workbook.setForceFormulaRecalculation(true);
            workbook.setActiveSheet(1);
            try (OutputStream os = new FileOutputStream("target/sxssf_stress2_output.xlsx")) {
                transformer.getWorkbook().write(os);
            }
        }
    }

    static class CellRefUpdater implements CellDataUpdater {
        
        @Override
        public void updateCellData(CellData cellData, CellRef targetCell, Context context) {
            if (cellData.isFormulaCell() && cellData.getFormula() != null) {
                cellData.setEvaluationResult(cellData.getFormula().replaceAll("(?<=[A-Za-z])\\d",
                        Integer.toString(targetCell.getRow() + 1)));
            }
        }
    }

    static class TotalCellUpdater implements CellDataUpdater {
        
        @Override
        public void updateCellData(CellData cellData, CellRef targetCell, Context context) {
            if (cellData.isFormulaCell() && cellData.getFormula().equals("SUM(E2)")) {
                String resultFormula = String.format("SUM(E2:E%d)", targetCell.getRow());
                cellData.setEvaluationResult(resultFormula);
            }
        }
    }

    static class DepartmentTotalCellUpdater implements CellDataUpdater {
        
        @Override
        public void updateCellData(CellData cellData, CellRef targetCell, Context context) {
            if (cellData.isFormulaCell() && cellData.getFormula().equals("SUM(E5)")) {
                Department department = (Department) context.getVar("department");
                if (department != null) {
                    String resultFormula = String.format("SUM(E%d:E%d)",
                            targetCell.getRow() - department.getStaff().size() + 1, targetCell.getRow());
                    cellData.setEvaluationResult(resultFormula);
                }
            }
        }
    }
}
