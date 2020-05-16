package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * Wrong cell ref replacement
 */
public class IssueB122Test {

    @Test
    public void test() throws IOException {
        // Prepare
        List<Person> persons = new ArrayList<>();
        persons.add(new Person(1, "Florian"));
        persons.add(new Person(50, "Michael"));
        persons.add(new Person(55, "Peter"));
        persons.add(new Person(40, "Stefan"));
        persons.add(new Person(45, "Marcus"));
        persons.add(new Person(49, "Thomas"));
        Context context = new Context();
        context.putVar("persons", persons);

        // Test
        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);

        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet("Tabelle1");
            assertEquals(50d, w.getCellValueAsDouble(4, 38), 0.1d);
            assertEquals(50d, w.getCellValueAsDouble(4, 39), 0.1d); // column AM
            int row = 12;
            for (Person person : persons) {
                assertEquals(Double.valueOf(person.getMerkmal()), w.getCellValueAsDouble(row++, 39), 0.1d);
            }
        }
    }

    public static class Person {
        private final String vorname; // forename
        private Integer merkmal;      // attribute

        public Person(int merkmal, String vorname) {
            this.vorname = vorname;
            this.merkmal = Integer.valueOf(merkmal);
        }

        public String getVorname() {
            return vorname;
        }

        public Integer getMerkmal() {
            return merkmal;
        }
    }
}
