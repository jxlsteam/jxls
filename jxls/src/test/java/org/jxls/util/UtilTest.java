package org.jxls.util;

import static java.util.Arrays.asList;
import static java.util.Collections.singletonList;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.jxls.util.Util.getSheetsNameOfMultiSheetTemplate;

import java.lang.reflect.InvocationTargetException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.beanutils.BasicDynaClass;
import org.apache.commons.beanutils.DynaBean;
import org.apache.commons.beanutils.DynaClass;
import org.apache.commons.beanutils.DynaProperty;
import org.junit.Test;
import org.jxls.area.Area;
import org.jxls.area.XlsArea;
import org.jxls.command.EachCommand;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.expression.Dummy;
import org.jxls.expression.JexlExpressionEvaluator;

public class UtilTest {
    // see more tests in UtilCreateTargetCellRefTest

    @Test
    public void should_return_sheet_names_of_multi_sheet_template() {
        // GIVEN
        EachCommand eachCommand = new EachCommand();
        eachCommand.setMultisheet("multiSheetOutputNames");

        Area areaWithMultiSheetOutputCommand = new XlsArea(new CellRef("areaWithMultiSheetOutput", 1, 1), new Size(1, 1));
        AreaRef ref = new AreaRef(new CellRef("areaWithMultiSheetOutput", 1, 1), new CellRef("areaWithMultiSheetOutput", 1, 1));
        areaWithMultiSheetOutputCommand.addCommand(ref, eachCommand);

        Area areaWithoutMultiSheetOutputCommand = new XlsArea(new CellRef("areaWithoutMultiSheetOutput", 1, 1), new Size());

        // WHEN
        List<String> sheetsNameOfMultiSheetTemplate = getSheetsNameOfMultiSheetTemplate(asList(areaWithMultiSheetOutputCommand, areaWithoutMultiSheetOutputCommand));

        // THEN
        assertEquals(sheetsNameOfMultiSheetTemplate, singletonList("areaWithMultiSheetOutput"));
    }
    
    @Test
    public void getObjectProperty_Map() throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        // Prepare
        Map<String, String> map = new HashMap<>();
        map.put("foo", "bar");
        
        // Test
        String r = (String) Util.getObjectProperty(map, "foo");
        
        // Verify
        assertEquals("bar", r);
    }
    
    @Test
    public void getObjectProperty_DynaBean() throws IllegalAccessException, InstantiationException, NoSuchMethodException, InvocationTargetException {
        // Prepare
        DynaClass dynaClass = new BasicDynaClass("Employee", null,
                new DynaProperty[] { new DynaProperty("name", String.class), });
        DynaBean bond = dynaClass.newInstance();
        bond.set("name", "James Bond 007");

        // Test
        String r = (String) Util.getObjectProperty(bond, "name");
        
        // Verify
        assertEquals("James Bond 007", r);
    }
    
    @Test
    public void getObjectProperty_JavaBean() throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        // Prepare
        Person bond = new Person("James Bond", 42, "London");
        bond.setDummy(new Dummy("007")); // nested attribute

        // Test
        String name = (String) Util.getObjectProperty(bond, "name");
        String number = (String) Util.getObjectProperty(bond, "dummy.strValue");
        
        // Verify
        assertEquals("James Bond", name);
        assertEquals("007", number);
    }
    
    /** Return empty collection instead of throwing exception if EachCommand.items resolves to null. */
    @Test
    public void issue200() {
        // Prepare
        JexlExpressionEvaluator anyEvaluator = new JexlExpressionEvaluator();
        Context emptyContext = new Context();
        
        // Test
        Iterable<Object> ret = Util.transformToIterableObject(anyEvaluator, "notExisting", emptyContext);
        
        // Verify
        assertFalse("Collection must be empty", ret.iterator().hasNext());
    }
}
