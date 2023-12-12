package org.jxls.examples.stress; // stress subpackage: long lasting tests

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
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn
 */
@Ignore
public class StressDemo {
    private static final Logger logger = LoggerFactory.getLogger(StressDemo.class);
    private static final int EMPLOYEE_COUNT = 30000;
    private static final int DEPARTMENT_COUNT = 100;
    private static final int DEP_EMPLOYEE_COUNT = 500;

    @Test
    public void test() throws IOException, InvalidFormatException {
        logger.info("Entering Stress demo");
        executeStress1();
        executeStress2();
    }

    private void executeStress1() throws IOException, InvalidFormatException {
        logger.info("Running Stress demo 1");
        logger.info("Generating " + EMPLOYEE_COUNT + " employees..");
        List<Employee> employees = Employee.generate(EMPLOYEE_COUNT);
        logger.info("Created " + employees.size() + " employees");
        try (InputStream is = StressDemo.class.getResourceAsStream("stress1.xls")) {
            try (OutputStream os = new FileOutputStream("target/stress1_output.xls")) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
                List<Area> xlsAreaList = areaBuilder.build(transformer, true);
                Area xlsArea = xlsAreaList.get(0);
                Context context = PoiTransformer.createInitialContext();
                context.putVar("employees", employees);
                long startTime = System.nanoTime();
                xlsArea.applyAt(new CellRef("Sheet2!A1"), context);
                xlsArea.processFormulas();
                long endTime = System.nanoTime();
                logger.info("Stress demo 1 time (s): " + (endTime - startTime) / 1000000000);
                transformer.write();
            }
        }
    }

    private void executeStress2() throws IOException, InvalidFormatException {
        logger.info("Running Stress demo 2");
        logger.info(
                "Generating " + DEPARTMENT_COUNT + " departments with " + DEP_EMPLOYEE_COUNT + " employees in each");
        List<Department> departments = Department.generate(DEPARTMENT_COUNT, DEP_EMPLOYEE_COUNT);
        logger.info("Created " + departments.size() + " departments");
        try (InputStream is = StressDemo.class.getResourceAsStream("stress2.xls")) {
            try (OutputStream os = new FileOutputStream("target/stress2_output.xls")) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
                List<Area> xlsAreaList = areaBuilder.build(transformer, true);
                Area xlsArea = xlsAreaList.get(0);
                Context context = PoiTransformer.createInitialContext();
                context.putVar("departments", departments);
                long startTime = System.nanoTime();
                xlsArea.applyAt(new CellRef("Sheet2!A1"), context);
                xlsArea.processFormulas();
                long endTime = System.nanoTime();
                logger.info("Stress demo 2 time (s): " + (endTime - startTime) / 1000000000);
                transformer.write();
            }
        }
    }
}
