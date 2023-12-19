package org.jxls.command;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.mockito.Mockito.when;

import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.jxls.area.Area;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.TransformationConfig;
import org.jxls.transform.Transformer;
import org.mockito.Mock;
import org.mockito.junit.MockitoJUnitRunner;

@RunWith(MockitoJUnitRunner.class)
public class EachCommandTest {
    @Mock
    private Area area;
    @Mock
    private Transformer transformer;
    @Mock
    private TransformationConfig transformationConfig;
    @Mock
    private ExpressionEvaluator expressionEvaluator;
    @Mock
    private Size size;

    @Test
    public void shouldRestoreVarNameInContextIfExists() {
        // given
        CellRef cellRef = new CellRef("A1");
        Context context = new Context();
        context.putVar("myVar", "Value 1");
        List<Object> list = new ArrayList<>();
        list.add("item 1");
        list.add("item 2");
        list.add("item 3");
        context.putVar("items", list);
        EachCommand eachCommand = new EachCommand("myVar", "items", area);
        when(area.getTransformer()).thenReturn(transformer);
        when(transformer.getTransformationConfig()).thenReturn(transformationConfig);
        when(eachCommand.getExpressionEvaluator()).thenReturn(expressionEvaluator);
        when(eachCommand.transformToIterableObject("items", context)).thenReturn(list);
        when(area.applyAt(cellRef, context)).thenReturn(size);
        // when
        eachCommand.applyAt(cellRef, context);
        // then
        assertEquals("original value for loop variable is not restored!", "Value 1", context.getVar("myVar"));
    }

    @Test
    public void shouldIgnoreMissingItems() {
        // given
        CellRef cellRef = new CellRef("A1");
        Context context = new Context();
        EachCommand eachCommand = new EachCommand("myVar", "items", area);
        when(area.getTransformer()).thenReturn(transformer);
        when(transformer.getTransformationConfig()).thenReturn(transformationConfig);
        when(transformer.getTransformationConfig().getExpressionEvaluator()).thenReturn(new JexlExpressionEvaluator());
        // when
        Size size = eachCommand.applyAt(cellRef, context);
        // then
        assertNotNull(size);
    }

    /** Return empty collection instead of throwing exception if EachCommand.items resolves to null. */
    @Test
    public void issue200() {
        // Prepare
        Context emptyContext = new Context();
        EachCommand each = new EachCommand() {
            @Override
            protected ExpressionEvaluator getExpressionEvaluator() {
                return new JexlExpressionEvaluator();
            }
        };
        
        // Test
        Iterable<Object> ret = each.transformToIterableObject("notExisting", emptyContext);
        
        // Verify
        assertFalse("Collection must be empty", ret.iterator().hasNext());
    }
}
