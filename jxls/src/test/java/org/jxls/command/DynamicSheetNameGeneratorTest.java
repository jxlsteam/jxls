package org.jxls.command;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

import java.util.Map;

import org.junit.Test;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.expression.ExpressionEvaluator;

public class DynamicSheetNameGeneratorTest {

    @Test
    public void test() {
        // Prepare
        Context context = new Context();
        context.putVar("sheetnames", "doe");
        
        // Test
        DynamicSheetNameGenerator gen = new DynamicSheetNameGenerator("sheetnames", new CellRef("A1"), new TestExpressionEvaluator());
        
        // Verify
        assertEquals("doe!A1", gen.generateCellRef(0, context).toString());
        assertEquals("'doe(1)'!A1", gen.generateCellRef(0 /* arg not used */, context).toString());
        assertEquals("'doe(2)'!A1", gen.generateCellRef(1, context).toString());
    }

    @Test
    public void testNull() {
        DynamicSheetNameGenerator gen = new DynamicSheetNameGenerator("N/A", new CellRef("A1"), new TestExpressionEvaluator());
        assertNull(gen.generateCellRef(0, new Context()));
    }
    
    private static class TestExpressionEvaluator implements ExpressionEvaluator {
        
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
