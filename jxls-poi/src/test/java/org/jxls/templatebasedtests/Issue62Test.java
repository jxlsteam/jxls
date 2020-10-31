package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * Using "each" and "merge" cause problem
 */
public class Issue62Test {

    @Test
    public void test() throws IOException {
        // Prepare
        List<TestBean> list = new ArrayList<>();
        List<TupleBean> tbList1 = new ArrayList<>();
        Collections.addAll(
                tbList1,
                new TupleBean("name11", "value11"),
                new TupleBean("name12", "value12")
        );
        List<TupleBean> tbList2 = new ArrayList<>();
        Collections.addAll(
                tbList2,
                new TupleBean("name21", "value21"),
                new TupleBean("name22", "value22")
        );
        Collections.addAll(
                list,
                new TestBean("title1", tbList1, "other11", "other12"),
                new TestBean("title2", tbList2, "other21", "other22")
        );
        Context context = new Context();
        context.putVar("list", list);
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            assertEquals("other12", w.getCellValueAsString(13, /*H*/8));
            assertEquals("There must not be empty rows between the tables!\n", "ov1", w.getCellValueAsString(16, /*G*/7));
            assertEquals("other21", w.getCellValueAsString(16, /*H*/8));
            assertEquals("other22", w.getCellValueAsString(17, /*H*/8));
        }
    }

    public static class TestBean {
        private final String title;
        private final List<TupleBean> tupleBeanList;
        private final String otherV1;
        private final String otherV2;

        public TestBean(String title, List<TupleBean> tupleBeanList, String otherV1, String otherV2) {
            this.title = title;
            this.tupleBeanList = tupleBeanList;
            this.otherV1 = otherV1;
            this.otherV2 = otherV2;
        }

        public String getTitle() {
            return title;
        }

        public List<TupleBean> getTupleBeanList() {
            return tupleBeanList;
        }

        public String getOtherV1() {
            return otherV1;
        }

        public String getOtherV2() {
            return otherV2;
        }
    }

    public static class TupleBean {
        private final String name;
        private final String value;

        public TupleBean(String name, String value) {
            this.name = name;
            this.value = value;
        }

        public String getName() {
            return name;
        }

        public String getValue() {
            return value;
        }
    }
}
