package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

public class IssueB206Test {
    // worked with 2.6.0, failure in 2.8.0-rc1

    @Test
    public void empty() throws Exception {
        // Test with empty lists
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(new Context());

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("TransactionAdditionalInfo");
            assertEquals("Category", w.getCellValueAsString(1, 1));
            assertEquals("Sub category", w.getCellValueAsString(1, 2)); // error
        }
    }

    @Test
    public void filled() throws Exception {
        // Prepare
        Context context = new Context();
        context.putVar("lines", createLines());
        context.putVar("titles", createTitles());
        context.putVar("subcatTitles", createSubcatTitles());
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("TransactionAdditionalInfo");
            assertEquals("Category", w.getCellValueAsString(1, 1));
            assertEquals("A", w.getCellValueAsString(2, 1));
            assertEquals("T-1", w.getCellValueAsString(1, 2));
            assertEquals("V-1", w.getCellValueAsString(2, 2));
            assertEquals("SCT-2", w.getCellValueAsString(1, 6));
            assertEquals("V-2", w.getCellValueAsString(2, 6));
            assertEquals("Sub category", w.getCellValueAsString(1, 4)); // error
            assertEquals("A-1", w.getCellValueAsString(2, 4)); // error
        }
    }

    private List<Line> createLines() {
        List<Line> ret = new ArrayList<>();
        Line line = new Line();
        line.setCategory("A");
        line.setValues(createValues());
        line.setSubcatTitle("A-1");
        line.setSubcatValues(createValues());
        ret.add(line);
        return ret;
    }

    private List<String> createValues() {
        List<String> ret = new ArrayList<>();
        ret.add("V-1");
        ret.add("V-2");
        return ret;
    }

    private List<String> createTitles() {
        List<String> ret = new ArrayList<>();
        ret.add("T-1");
        ret.add("T-2");
        return ret;
    }

    private List<String> createSubcatTitles() {
        List<String> ret = new ArrayList<>();
        ret.add("SCT-1");
        ret.add("SCT-2");
        return ret;
    }

    public static class Line {
        private String category;
        private String subcatTitle;
        private List<String> values;
        private List<String> subcatValues;

        public String getCategory() {
            return category;
        }

        public void setCategory(String category) {
            this.category = category;
        }

        public String getSubcatTitle() {
            return subcatTitle;
        }

        public void setSubcatTitle(String subcatTitle) {
            this.subcatTitle = subcatTitle;
        }

        public List<String> getValues() {
            return values;
        }

        public void setValues(List<String> values) {
            this.values = values;
        }

        public List<String> getSubcatValues() {
            return subcatValues;
        }

        public void setSubcatValues(List<String> subcatValues) {
            this.subcatValues = subcatValues;
        }
    }
}
