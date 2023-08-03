package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import org.apache.poi.ss.usermodel.Cell;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.templatebasedtests.Issue17Test.Item;
import org.jxls.transform.poi.PoiUtil;
import org.jxls.transform.poi.ValueWithCellAccess;

/**
 * This testcase writes for every cell value its origin as a Excel cell comment.
 */
public class ValueWithCellAccessTest {
    
    @Test
    public void test() throws Exception {
        // Prepare
        Context context = new Context();
        context.putVar("items", Issue17Test.getItems());
        context.putVar("_f", new Functions(context)); // withSource() method returns ValueWithCellAccess
        
        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            // cell values:
            assertEquals("ANDROID", w.getCellValueAsString(2, 2));
            assertEquals("JP", w.getCellValueAsString(2, 3));
            // origins as comments:
            assertEquals("class=org.jxls.templatebasedtests.Issue17Test$Item\nfield=app_type\nid=registered_user", w.getCellComment(2, 2));
            assertEquals("class=org.jxls.templatebasedtests.Issue17Test$Item\nfield=country_code\nid=registered_user", w.getCellComment(2, 3));
        }
    }
    
    public static class Functions {
        private final Context context;
        
        public Functions(Context ctx) {
            this.context = ctx;
        }
        
        public ValueWithCellAccess withSource(String path) {
            if (path == null || !path.contains(".")) {
                throw new RuntimeException("path must not be null and must contain '.'");
            }
            int o = path.lastIndexOf(".");
            String head = path.substring(0, o);
            String tail = path.substring(o + 1); // field name
            Item pe = (Item) evaluate(head);
            Object value;
            switch (tail) {
            case "app_type":
                value = pe.getApp_type();
                break;
            case "country_code":
                value = pe.getCountry_code();
                break;
            default:
                throw new RuntimeException("Unknown field: " + tail);
            }
            return new MyValueWithCellAccess(pe, tail, value);
        }

        private Object evaluate(String expression) {
            return new JexlExpressionEvaluator(expression).evaluate(context.toMap());
        }
    }

    public static class MyValueWithCellAccess implements ValueWithCellAccess {
        private Item pe;
        private String fieldName;
        private Object value;

        public MyValueWithCellAccess(Item pe, String fieldName, Object value) {
            this.pe = pe;
            this.fieldName = fieldName;
            this.value = value;
        }
        
        @Override
        public Object getValue() {
            return value;
        }

        @Override
        public void applyBefore(Cell cell) {
            String note = "class=" + pe.getClass().getName() + "\nfield=" + fieldName + "\nid=" + pe.getId();
            PoiUtil.setCellComment(cell, note, "auto", PoiUtil.createAnchor(cell, 1, 3, 0, 4));
        }

        @Override
        public void applyAfter(Cell cell) {
        }
    }
}
