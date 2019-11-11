package org.jxls.templatebasedtests.multisheet;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.jxls.expression.ExpressionEvaluator;

public abstract class AbstractMultiSheetTest {
    
    public static class TestSheet {
        private String name;
        private final List<String> items = new ArrayList<>();

        public TestSheet(String name) {
            this.name = name;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public List<String> getItems() {
            return items;
        }
    }

    protected List<TestSheet> getTestSheets() {
        List<TestSheet> testSheets = new ArrayList<>();
        TestSheet s = new TestSheet("data");
        s.getItems().add("d-1");
        s.getItems().add("d-2");
        s.getItems().add("d-3");
        testSheets.add(s);
        s = new TestSheet("parameters");
        s.getItems().add("p.A");
        s.getItems().add("p.B");
        testSheets.add(s);
        return testSheets;
    }

    protected List<String> getSheetnames(List<TestSheet> testSheets) {
        List<String> sheetnames = new ArrayList<>();
        for (TestSheet i : testSheets) {
            sheetnames.add(i.getName());
        }
        return sheetnames;
    }

    public static class TestExpressionEvaluator implements ExpressionEvaluator {
        
        @Override
        public Object evaluate(String expression, Map<String, Object> context) {
            return context.get(expression);
        }

        @Override
        public Object evaluate(Map<String, Object> context) {
            throw new UnsupportedOperationException();
        }

        @Override
        public String getExpression() {
            throw new UnsupportedOperationException();
        }
    }
}
