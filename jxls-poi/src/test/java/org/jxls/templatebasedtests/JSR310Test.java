package org.jxls.templatebasedtests;

import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Collections;

import static org.junit.Assert.assertEquals;

public class JSR310Test {

    private static final LocalDate SAMPLE_DATE = LocalDate.parse("2003-07-27");
    private static final LocalDateTime SAMPLE_DATETIME = LocalDateTime.parse("2020-02-25T15:07:43");

    private static TestWorkbook w;

    @BeforeClass
    public static void setUpClass() {
        // Prepare
        Context context = new Context();
        context.putVar("samples", Collections.singletonList(new Sample()));

        JxlsTester tester = JxlsTester.xlsx(JSR310Test.class);
        tester.processTemplate(context);

        w = tester.getWorkbook();
        w.selectSheet(0);
    }

    @AfterClass
    public static void tearDownClass() {
        w.close();
    }

    @Test
    public void testLocalDate() {
        assertEquals(SAMPLE_DATE.atStartOfDay(), w.getCellValueAsLocalDateTime(4, 1));
    }

    @Test
    public void testLocalDateTime() {
        assertEquals(SAMPLE_DATETIME, w.getCellValueAsLocalDateTime(4, 2));
    }

    @SuppressWarnings("unused")
    public static class Sample {
        public LocalDate localDate = SAMPLE_DATE;
        public LocalDateTime localDateTime = SAMPLE_DATETIME;
    }

}
