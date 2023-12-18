package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.entity.Employee;
import org.jxls.util.JxlsHelper;

/**
 * Created by Leonid Vysochyn on 30-Jun-15.
 */
public class GridCommandDemo {

    @Test
    public void test() throws ParseException, IOException {
        List<Employee> employees = Employee.generateSampleEmployeeData();
        executeGridMatrixDemo(employees);
        executeGridObjectListDemo(employees);
    }

    private void executeGridMatrixDemo(List<Employee> employees) throws IOException {
        List<List<Object>> data = createGridData(employees);
        try(InputStream is = GridCommandDemo.class.getResourceAsStream("grid_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/grid_output1.xls")) {
                Context context = new Context();
                context.putVar("headers", Arrays.asList("Name", "Birthday", "Payment"));
                context.putVar("data", data);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }

    private void executeGridObjectListDemo(List<Employee> employees) throws IOException {
        try (InputStream is = GridCommandDemo.class.getResourceAsStream("grid_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/grid_output2.xls")) {
                Context context = new Context();
                context.putVar("headers", Arrays.asList("Name", "Birthday", "Payment"));
                context.putVar("data", employees);
                JxlsHelper.getInstance().processGridTemplateAtCell(is, os, context, "name,birthDate,payment", "Sheet2!A1");
            }
        }
    }

    // also used by IssueB090Test
    public static List<List<Object>> createGridData(List<Employee> employees) {
        List<List<Object>> data = new ArrayList<>();
        for (Employee employee : employees) {
            data.add( convertEmployeeToList(employee));
        }
        return data;
    }

    private static List<Object> convertEmployeeToList(Employee employee) {
        List<Object> list = new ArrayList<>();
        list.add(employee.getName());
        list.add(employee.getBirthDate());
        list.add(employee.getPayment());
        return list;
    }
}
