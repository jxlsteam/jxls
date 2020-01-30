package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xml.XmlAreaBuilder;
import org.jxls.command.TestWorkbook;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;

/**
 * Tests groupBy and groupOrder for XML syntax
 */
public class Issue196Test {

    @Test
    public void test() throws IOException {
        // Prepare
        List<ClockInOut> clockInOuts = new ArrayList<>();
        clockInOuts.add(new ClockInOut("E00001", "Mayor, Tom", "Boss", "Ruler"));
        
        // Test
        try (InputStream is = Issue196Test.class.getResourceAsStream("Issue196Test.xlsx")) {
            try (OutputStream os = new FileOutputStream("target/Issue196Test_output.xlsx")) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                try (InputStream configInputStream = Issue196Test.class.getResourceAsStream("Issue196Test.xml")) {
                    AreaBuilder areaBuilder = new XmlAreaBuilder(configInputStream, transformer);
                    List<Area> xlsAreaList = areaBuilder.build();
                    Area xlsArea = xlsAreaList.get(0);
                    Context context = new Context();
                    context.putVar("clockinouts", clockInOuts);
                    xlsArea.applyAt(new CellRef("Template!A1"), context);
                    transformer.write();
                }
            }
        }

        // Verify
        try (TestWorkbook xls = new TestWorkbook(new File("target/Issue196Test_output.xlsx"))) {
            xls.selectSheet("Template");
            assertEquals("Mayor, Tom", xls.getCellValueAsString(2, 2));
        }
    }
    
    public static class ClockInOut {
        private final String employeeId;
        private final String employeeName;
        private final String employeeType;
        private final String assignment;

        public ClockInOut(String employeeId, String employeeName, String employeeType, String assignment) {
            this.employeeId = employeeId;
            this.employeeName = employeeName;
            this.employeeType = employeeType;
            this.assignment = assignment;
        }

        public String getEmployeeId() {
            return employeeId;
        }

        public String getEmployeeName() {
            return employeeName;
        }

        public String getEmployeeType() {
            return employeeType;
        }

        public String getAssignment() {
            return assignment;
        }
    }
}
