package org.jxls.command;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.functions.BigDecimalSummarizerBuilder;
import org.jxls.functions.DoubleSummarizerBuilder;
import org.jxls.functions.GroupSum;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.JxlsHelper;

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
    public void testWithMapsAndDouble() throws IOException, InvalidFormatException {
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
    public void testWithBeansAndBigDecimal() throws IOException, InvalidFormatException {
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
    
    private void check(Context context) throws IOException, InvalidFormatException {
        // Test
        InputStream in = GroupSumTest.class.getResourceAsStream("groupSum.xlsx");
        File outputFile = new File("target/groupSum.xlsx");
        FileOutputStream out = new FileOutputStream(outputFile);
        PoiTransformer transformer = PoiTransformer.createTransformer(in, out);
        JxlsHelper.getInstance().processTemplate(context, transformer);
        
        // Verify
        try (TestWorkbook xls = new TestWorkbook(outputFile)) {
            xls.selectSheet("Group sums");
            assertEquals("1st group sum is wrong! (Main department) E5\n", Double.valueOf(170000d), xls.getCellValueAsDouble(5, 5));
            assertEquals("2nd group sum is wrong! (Finance department) E10\n", Double.valueOf(130000d), xls.getCellValueAsDouble(10, 5));
            assertEquals("Total sum (calculated by fx:sum) in cell E12 is wrong!\n", Double.valueOf(300000d), xls.getCellValueAsDouble(12, 5));
        }
    }
}
