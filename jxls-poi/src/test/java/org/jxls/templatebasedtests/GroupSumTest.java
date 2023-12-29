package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.entity.Employee;
import org.jxls.functions.BigDecimalSummarizerBuilder;
import org.jxls.functions.DoubleSummarizerBuilder;
import org.jxls.functions.GroupSum;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

/**
 * Group sum test
 * 
 * <p>The test file contains a function with 2nd arg as String and also a function with a JEXL expression.</p>
 */
public class GroupSumTest {

    /**
     * This test uses Map objects. The salary is of type Double.
     */
    @Test
    public void testWithMapsAndDouble() {
        List<Map<String, Object>> maps = new ArrayList<>();
        maps.add(createEmployee("03 Finance department", "Christiane", "Operator", "Hartefeld", 40000));
        maps.add(createEmployee("01 Main department", "Claudia", "Assistent", "Issum", 30000));
        maps.add(createEmployee("03 Finance department", "Nadine", "Leader", "Mönchengladbach", 90000));
        maps.add(createEmployee("01 Main department", "Sven", "Mayor", "Veert", 140000));
        Map<String, Object> data = new HashMap<>();
        data.put("details", maps);
        check(data, new GroupSum<Double>(new DoubleSummarizerBuilder()));
    }

    private Map<String, Object> createEmployee(String department, String name, String job, String city, double salary) {
        Map<String, Object> map = new HashMap<>();
        map.put("buGroup", department);
        map.put("name", name);
        map.put("payment", Double.valueOf(salary));
        map.put("city", city);
        return map;
    }

    /**
     * This test uses beans. The salary is of type BigDecimal.
     */
    @Test
    public void testWithBeansAndBigDecimal() {
        List<Employee> beans = new ArrayList<>();
        beans.add(newEmployee("03 Finance department", "Christiane", "Operator", "Hartefeld", 40000));
        beans.add(newEmployee("01 Main department", "Claudia", "Assistent", "Issum", 30000));
        beans.add(newEmployee("03 Finance department", "Nadine", "Leader", "Mönchengladbach", 90000));
        beans.add(newEmployee("01 Main department", "Sven", "Mayor", "Veert", 140000));
        Map<String, Object> data = new HashMap<>();
        data.put("details", beans);
        check(data, new GroupSum<BigDecimal>(new BigDecimalSummarizerBuilder()));
    }
    
    private Employee newEmployee(String department, String name, String job, String city, double salary) {
        return new Employee(name, null, salary, 0, department);
    }
    
    private void check(Map<String, Object> data, GroupSum<?> groupSum) {
        data.put("G", groupSum);
        
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass());
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance().needsPublicContext(groupSum));
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Group sums");
            assertEquals("1st group sum is wrong! (Main department) E5\n", 170000d, w.getCellValueAsDouble(5, 5), 0.005d);
            assertEquals("2nd group sum is wrong! (Finance department) E10\n", 130000d, w.getCellValueAsDouble(10, 5), 0.005d);
            assertEquals("Total sum (calculated by G.sum) in cell E12 is wrong!\n", 300000d, w.getCellValueAsDouble(12, 5), 0.005d);
        }
    }
    
    @Test
    public void filterCondition() {
        // Prepare
        List<Map<String, Object>> maps = new ArrayList<>();
        maps.add(createEmployee("01 Main department", "Sven", "Mayor", "Geldern", 140000));
        maps.add(createEmployee("01 Main department", "Dagmar", "Assistent", "Geldern", 32000));
        maps.add(createEmployee("01 Main department", "Claudia", "Assistent", "Issum", 30000));
        maps.add(createEmployee("01 Main department", "Draci the dragon", "Mascot", "Wetten", -1));
        Map<String, Object> data = new HashMap<>();
        data.put("details", maps);
        GroupSum<Double> groupSum = new GroupSum<Double>(new DoubleSummarizerBuilder());
        data.put("G", groupSum);
        
        // Test
        Jxls3Tester tester = Jxls3Tester.xlsx(getClass(), "filterCondition");
        tester.test(data, JxlsPoiTemplateFillerBuilder.newInstance().needsPublicContext(groupSum));
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Group sums");
            assertEquals("wrong Geldern sum (E8)\n", 140000 + 32000, w.getCellValueAsDouble(8, 5), 0.005d);
            assertEquals("wrong outside Geldern sum (E9)\n", 30000, w.getCellValueAsDouble(9, 5), 0.005d);
        }
    }
}
