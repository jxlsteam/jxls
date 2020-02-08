package org.jxls.command;

import static org.junit.Assert.assertEquals;
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
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.TransformationConfig;
import org.jxls.transform.Transformer;
import org.jxls.util.UtilWrapper;
import org.mockito.Mock;
import org.mockito.junit.MockitoJUnitRunner;

@RunWith(MockitoJUnitRunner.class)
public class EachCommandJTest {
    @Mock
    private Area area;
    @Mock
    private Transformer transformer;
    @Mock
    private TransformationConfig transformationConfig;
    @Mock
    private Size size;
    @Mock
    private UtilWrapper util;

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
        eachCommand.setUtil(util);
        when(area.getTransformer()).thenReturn(transformer);
        when(transformer.getTransformationConfig()).thenReturn(transformationConfig);
        when(util.transformToIterableObject(null, "items", context)).thenReturn(list);
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
        eachCommand.setUtil(new UtilWrapper());
        when(area.getTransformer()).thenReturn(transformer);
        when(transformer.getTransformationConfig()).thenReturn(transformationConfig);
        when(transformer.getTransformationConfig().getExpressionEvaluator()).thenReturn(new JexlExpressionEvaluator());
        // when
        Size size = eachCommand.applyAt(cellRef, context);
        // then
        assertNotNull(size);
    }
}
