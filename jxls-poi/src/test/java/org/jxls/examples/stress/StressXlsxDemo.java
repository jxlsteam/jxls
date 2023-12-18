package org.jxls.examples.stress;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Ignore;
import org.junit.Test;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.entity.Department;
import org.jxls.entity.Employee;
import org.jxls.formula.FastFormulaProcessor;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiContext;
import org.jxls.transform.poi.PoiTransformer;

/**
 * @author Leonid Vysochyn
 */
@Ignore
public class StressXlsxDemo {
    public static final int EMPLOYEE_COUNT = 30000;
    public static final int DEPARTMENT_COUNT = 100;
    public static final int DEP_EMPLOYEE_COUNT = 500;

    @Test
    public void test() throws IOException, InvalidFormatException {
        executeStress1();
        executeStress2();
    }

    private void executeStress1() throws IOException {
        List<Employee> employees = Employee.generate(EMPLOYEE_COUNT);
        try (InputStream is = StressXlsxDemo.class.getResourceAsStream("stress1.xlsx")) {
            try (OutputStream os = new FileOutputStream("target/stress1_output.xlsx")) {
                Transformer transformer = PoiTransformer.createTransformer(is, os);
                AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
                List<Area> xlsAreaList = areaBuilder.build(transformer, true);
                Area xlsArea = xlsAreaList.get(0);
                Context context = new PoiContext();
                context.putVar("employees", employees);
                xlsArea.applyAt(new CellRef("Sheet2!A1"), context);
                xlsArea.setFormulaProcessor(new FastFormulaProcessor());
                xlsArea.processFormulas();
                transformer.write();
            }
        }
    }

    private void executeStress2() throws IOException {
        List<Department> departments = Department.generate(DEPARTMENT_COUNT, DEP_EMPLOYEE_COUNT);
        try (InputStream is = StressXlsxDemo.class.getResourceAsStream("stress2.xlsx")) {
            try (OutputStream os = new FileOutputStream("target/stress2_output.xlsx")) {
                Transformer transformer = PoiTransformer.createTransformer(is, os);
                AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
                List<Area> xlsAreaList = areaBuilder.build(transformer, true);
                Area xlsArea = xlsAreaList.get(0);
                Context context = new PoiContext();
                context.putVar("departments", departments);
                xlsArea.applyAt(new CellRef("Sheet2!A1"), context);
                xlsArea.setFormulaProcessor(new FastFormulaProcessor());
                xlsArea.processFormulas();
                transformer.write();
            }
        }
    }
}
