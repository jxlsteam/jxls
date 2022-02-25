package org.jxls.templatebasedtests;

import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

import java.time.*;
import java.time.temporal.ChronoUnit;
import java.util.Collections;

import static org.junit.Assert.assertEquals;

public class JSR310Test {

    private static final LocalDate SAMPLE_DATE = LocalDate.parse("2003-07-27");
    private static final LocalDateTime SAMPLE_DATETIME = LocalDateTime.parse("2020-02-25T15:07:43");
    private static final ZonedDateTime SAMPLE_ZONED_DATETIME = ZonedDateTime.parse("2005-05-20T13:31:34-03:00[America/Maceio]");
    private static final LocalTime SAMPLE_TIME = LocalTime.parse("09:45:00");
    private static final Instant SAMPLE_INSTANT = Instant.now();

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

    @Test
    public void testZonedLocalDateTime() {
        assertEquals(SAMPLE_ZONED_DATETIME.toLocalDateTime(), w.getCellValueAsLocalDateTime(4, 3));
    }

    @Test
    public void testLocalTime() {
        assertEquals(SAMPLE_TIME, w.getCellValueAsLocalDateTime(4, 4).toLocalTime());
    }

    @Test
    public void testInstant() {
        assertEquals(
                SAMPLE_INSTANT.atZone(ZoneId.systemDefault()).truncatedTo(ChronoUnit.MILLIS),
                w.getCellValueAsLocalDateTime(4, 5).atZone(ZoneId.systemDefault()).truncatedTo(ChronoUnit.MILLIS)
        );
    }

    @SuppressWarnings("unused")
    public static class Sample {
        public LocalDate localDate = SAMPLE_DATE;
        public LocalDateTime localDateTime = SAMPLE_DATETIME;
        public ZonedDateTime zonedDateTime = SAMPLE_ZONED_DATETIME;
        public LocalTime localTime = SAMPLE_TIME;
        public Instant instant = SAMPLE_INSTANT;
    }

}
