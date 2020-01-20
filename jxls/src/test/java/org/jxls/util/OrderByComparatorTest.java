package org.jxls.util;

import static org.junit.Assert.assertEquals;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.junit.Test;

public class OrderByComparatorTest {

    @Test
    public void test() {
        List<Person> persons = new ArrayList<>();
        persons.add(new Person("Maria", 20, "Berlin"));
        persons.add(new Person("Anna", 20, "Berlin"));
        persons.add(new Person(null, 20, "Berlin"));
        persons.add(new Person("Cloe", 20, "Stuttgart"));
        persons.add(new Person("Anna", 20, "Stuttgart"));
        persons.add(new Person("Jack", 20, null));
        
        List<String> orderBy = new ArrayList<>();
        orderBy.add("city desc");
        orderBy.add(" name");
        
        Collections.sort(persons, new OrderByComparator<Person>(orderBy, new UtilWrapper()));
        
        int i = 0;
        assertEquals("Jack/null", persons.get(i++).getNameAndCity());
        assertEquals("Anna/Stuttgart", persons.get(i++).getNameAndCity());
        assertEquals("Cloe/Stuttgart", persons.get(i++).getNameAndCity());
        assertEquals("Anna/Berlin", persons.get(i++).getNameAndCity());
        assertEquals("Maria/Berlin", persons.get(i++).getNameAndCity());
        assertEquals("null/Berlin", persons.get(i++).getNameAndCity());
    }
}
