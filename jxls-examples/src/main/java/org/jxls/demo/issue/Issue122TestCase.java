package org.jxls.demo.issue;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class Issue122TestCase {
    public static void main(String[] args) throws IOException {
        List<Issue122TestCase.Person> persons = new ArrayList<>();
        persons.add(new Issue122TestCase.Person(1, "Florian"));
        persons.add(new Issue122TestCase.Person(50, "Michael"));
        persons.add(new Issue122TestCase.Person(55, "Peter"));
        persons.add(new Issue122TestCase.Person(40, "Stefan"));
        persons.add(new Issue122TestCase.Person(45, "Marcus"));
        persons.add(new Issue122TestCase.Person(49, "Thomas"));

        try(InputStream is = Issue122TestCase.class.getResourceAsStream("issue122_template.xlsx")) {
            try (OutputStream os = new FileOutputStream("target/issue122_output.xlsx")) {
                Context context = new Context();
                context.putVar("persons", persons);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }

    public static class Person {
        private final String vorname;
        private Integer merkmal;

        public Person(String vorname, Integer merkmal) {
            this.vorname = vorname;
            this.merkmal = merkmal;
        }

        public Person(int merkmal, String vorname) {
            this(vorname, Integer.valueOf(merkmal));
        }

        public String getVorname() {
            return vorname;
        }

        public Integer getMerkmal() {
            return merkmal;
        }
    }
}
