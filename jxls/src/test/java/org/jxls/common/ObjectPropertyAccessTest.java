package org.jxls.common;

import static org.junit.Assert.assertEquals;

import java.lang.reflect.InvocationTargetException;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.beanutils2.BasicDynaClass;
import org.apache.commons.beanutils2.DynaBean;
import org.apache.commons.beanutils2.DynaClass;
import org.apache.commons.beanutils2.DynaProperty;
import org.junit.Test;
import org.jxls.command.Person;
import org.jxls.expression.Dummy;

public class ObjectPropertyAccessTest {
    
    @Test
    public void getObjectProperty_Map() throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        // Prepare
        Map<String, String> map = new HashMap<>();
        map.put("foo", "bar");
        
        // Test
        String r = (String) ObjectPropertyAccess.getObjectProperty(map, "foo");
        
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
        String r = (String) ObjectPropertyAccess.getObjectProperty(bond, "name");
        
        // Verify
        assertEquals("James Bond 007", r);
    }

    @Test
    public void getObjectProperty_JavaBean() throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        // Prepare
        Person bond = new Person("James Bond", 42, "London");
        bond.setDummy(new Dummy("007")); // nested attribute

        // Test
        String name = (String) ObjectPropertyAccess.getObjectProperty(bond, "name");
        String number = (String) ObjectPropertyAccess.getObjectProperty(bond, "dummy.strValue");
        
        // Verify
        assertEquals("James Bond", name);
        assertEquals("007", number);
    }

}
