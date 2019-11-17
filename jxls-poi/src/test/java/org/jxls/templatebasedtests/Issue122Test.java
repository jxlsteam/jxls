package org.jxls.templatebasedtests;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.common.Context;

/**
 * Wrong cell ref replacement
 */
public class Issue122Test {

    @Test
    public void test() throws IOException {
        List<Person> persons = new ArrayList<>();
        persons.add(new Person(1, "Florian"));
        persons.add(new Person(50, "Michael"));
        persons.add(new Person(55, "Peter"));
        persons.add(new Person(40, "Stefan"));
        persons.add(new Person(45, "Marcus"));
        persons.add(new Person(49, "Thomas"));
        Context context = new Context();
        context.putVar("persons", persons);

        JxlsTester tester = JxlsTester.xlsx(getClass());
        tester.processTemplate(context);
        
        // TODO Check testcase!
        // TODO assertions
    }

    public static class Person {
        private final String vorname;
        private Integer merkmal;

        public Person(String vorname, Integer merkmal) {
            this.vorname = vorname;
            this.merkmal = merkmal;
        }

        public Person(int merkmal, String vorname) {
            this(vorname, merkmal);
        }

        public String getVorname() {
            return vorname;
        }

        public Integer getMerkmal() {
            return merkmal;
        }
    }
}
