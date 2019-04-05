package org.jxls.util;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;

import org.junit.Test;
import org.jxls.common.GroupData;

/**
 * Test groupOrder algorithm. Used by EachCommand (groupOrder attribute).
 */
public class GroupOrderTest {

    /**
     * There's a difference between no groupOrder attribute and groupOrder="asc"!
     */
    @Test
    public void groupOrderNull() {
        Collection<GroupData> result = Util.groupIterable(getTestData(), "age", null);
        
        Iterator<GroupData> iter = result.iterator();
        assertEquals(25, nextAge(iter));
        assertEquals(20, nextAge(iter));
        assertEquals(30, nextAge(iter));
        assertFalse(iter.hasNext());
    }

    /**
     * If you use groupBy you should usually also use groupOrder="asc"!
     */
    @Test
    public void groupOrderAsc() {
        Collection<GroupData> result = Util.groupIterable(getTestData(), "age", "ASC");
        
        Iterator<GroupData> iter = result.iterator();
        assertEquals(20, nextAge(iter));
        assertEquals(25, nextAge(iter));
        assertEquals(30, nextAge(iter));
        assertFalse(iter.hasNext());
    }

    /**
     * Use groupOrder="desc" to reverse the group order. The case of "asc" and "desc" does not matter.
     */
    @Test
    public void groupOrderDesc() {
        Collection<GroupData> result = Util.groupIterable(getTestData(), "age", "desc");
        
        Iterator<GroupData> iter = result.iterator();
        assertEquals(30, nextAge(iter));
        assertEquals(25, nextAge(iter));
        assertEquals(20, nextAge(iter));
        assertFalse(iter.hasNext());
    }

    private List<Person> getTestData() {
        List<Person> persons = new ArrayList<>();
        persons.add(new Person("B", 25, "NY")); // first in the list
        persons.add(new Person("A", 20, "London"));
        persons.add(new Person("C", 20, "Paris"));
        persons.add(new Person("D", 30, "Stockholm"));
        persons.add(new Person("H", 25, "Paris"));
        return persons;
    }

    private int nextAge(Iterator<GroupData> iter) {
        return ((Person) iter.next().getItem()).getAge();
    }
}
