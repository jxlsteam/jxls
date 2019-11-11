package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.entity.TestEmployee;
import org.jxls.functions.BigDecimalSummarizerBuilder;
import org.jxls.functions.DoubleSummarizerBuilder;
import org.jxls.functions.GroupSum;

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
        Context context = new Context();
        context.putVar("details", maps);
        context.putVar("G", new GroupSum<Double>(context, new DoubleSummarizerBuilder()));
        check(context);
    }

    private Map<String, Object> createEmployee(String department, String name, String job, String city, double salary) {
        Map<String, Object> map = new HashMap<>();
        map.put("department", department);
        map.put("name", name);
        map.put("job", job);
        map.put("city", city);
        map.put("salary", Double.valueOf(salary));
        return map;
    }

    /**
     * This test uses beans. The salary is of type BigDecimal.
     */
    @Test
    public void testWithBeansAndBigDecimal() {
        List<TestEmployee> beans = new ArrayList<>();
        beans.add(new TestEmployee("03 Finance department", "Christiane", "Operator", "Hartefeld", 40000));
        beans.add(new TestEmployee("01 Main department", "Claudia", "Assistent", "Issum", 30000));
        beans.add(new TestEmployee("03 Finance department", "Nadine", "Leader", "Mönchengladbach", 90000));
        beans.add(new TestEmployee("01 Main department", "Sven", "Mayor", "Veert", 140000));
        Context context = new Context();
        context.putVar("details", beans);
        context.putVar("G", new GroupSum<BigDecimal>(context, new BigDecimalSummarizerBuilder()));
        check(context);
    }
    
    private void check(Context context) {
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Group sums");
            assertEquals("1st group sum is wrong! (Main department) E5\n", Double.valueOf(170000d), w.getCellValueAsDouble(5, 5));
            assertEquals("2nd group sum is wrong! (Finance department) E10\n", Double.valueOf(130000d), w.getCellValueAsDouble(10, 5));
            assertEquals("Total sum (calculated by fx:sum) in cell E12 is wrong!\n", Double.valueOf(300000d), w.getCellValueAsDouble(12, 5));
        }
    }
}
