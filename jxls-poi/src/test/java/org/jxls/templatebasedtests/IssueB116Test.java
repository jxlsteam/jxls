package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.util.Arrays;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * Formula handling issues (formula external to any jx:area)
 */
public class IssueB116Test {

    @Test
    public void externalFormulas() throws IOException {
        // Prepare
        Context context = getContext();
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass(), "externalFormulas");
        tester.setUseFastFormulaProcessor(false); // true TODO It looks like that the FastFormulaProcessor cannot handle this Excel template anymore. Leonid and Marcus set this to false because this testcase blocks the development.
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("test");
            assertEquals(1.234, w.getCellValueAsDouble(3, 2), 0.00001d);
            assertEquals(5.678, w.getCellValueAsDouble(5, 2), 0.00001d);
            assertEquals(3.1234, w.getCellValueAsDouble(7, 2), 0.00001d);
            assertEquals(8.909, w.getCellValueAsDouble(9, 2), 0.00001d);
            assertEquals(12.34567, w.getCellValueAsDouble(11, 2), 0.000001d);
            final int f = 6, g = 7, j = 10;
            assertEquals(10d, w.getCellValueAsDouble(3, f), 0.1d);
            assertEquals(10d, w.getCellValueAsDouble(3, j), 0.1d);
            assertEquals(10d, w.getCellValueAsDouble(4, f), 0.1d);
            assertEquals(10d, w.getCellValueAsDouble(4, g), 0.1d);
            assertEquals(10d, w.getCellValueAsDouble(3, f), 0.1d);
            assertEquals(10d, w.getCellValueAsDouble(3, f), 0.1d);
            assertEquals(10d, w.getCellValueAsDouble(3, f), 0.1d);
            assertEquals(10d, w.getCellValueAsDouble(3, f), 0.1d);
            assertEquals(5d, w.getCellValueAsDouble(8, f), 0.1d);
            assertEquals(5d, w.getCellValueAsDouble(8, g), 0.1d);
            assertEquals(5d, w.getCellValueAsDouble(8, j), 0.1d);
            assertEquals(5d, w.getCellValueAsDouble(9, f), 0.1d);
            assertEquals(5d, w.getCellValueAsDouble(9, g), 0.1d);
        }
    }

    @Test
    public void forEachFormulas() throws IOException {
        // Prepare
        Context context = getContext();

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass(), "forEachFormulas");
        tester.setUseFastFormulaProcessor(false); // false (=default value) was set in old Demo class
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("test");
            assertEquals(1.234, w.getCellValueAsDouble(3, 2), 0.00001d);
            assertEquals(5.678, w.getCellValueAsDouble(4, 2), 0.00001d);
            assertEquals(3.1234, w.getCellValueAsDouble(5, 2), 0.00001d);
            assertEquals(8.909, w.getCellValueAsDouble(6, 2), 0.00001d);
            assertEquals(12.34567, w.getCellValueAsDouble(7 , 2), 0.000001d);
            
            assertEquals(1.4808, w.getCellValueAsDouble(3, 3), 0.00001d);
            assertEquals(6.8136, w.getCellValueAsDouble(4, 3), 0.00001d);
            assertEquals(3.74808, w.getCellValueAsDouble(5, 3), 0.00001d);
            assertEquals(10.6908, w.getCellValueAsDouble(6, 3), 0.00001d);
            assertEquals(14.814804, w.getCellValueAsDouble(7 , 3), 0.00000001d);

            assertEquals(1.33302663139189, w.getCellValueAsDouble(3, 4), 0.00000000000001d);
            assertEquals(2.85942651592937, w.getCellValueAsDouble(4, 4), 0.00000000000001d);
            assertEquals(2.12077721602247, w.getCellValueAsDouble(5, 4), 0.00000000000001d);
            assertEquals(3.58175376038052, w.getCellValueAsDouble(6, 4), 0.00000000000001d);
            assertEquals(4.21636867458243, w.getCellValueAsDouble(7, 4), 0.00000000000001d);
        }
    }

    private Context getContext() {
        Context context = new Context();
        context.putVar("vars", Arrays.asList(1.234, 5.678, 3.1234, 8.9090, 12.34567));
        return context;
    }
}
