Object Collection Sample
========================

Introduction
-------------

This sample shows how to output a collection of Java objects into Excel with Jxls.

We will use a list of the following *Employee* objects to demonstrate how to output an object collection to Excel with Jxls.

    public class Employee {
        private String name;
        private int age;
        private Double payment;
        private Double bonus;
        private Date birthDate;
        private Employee superior;

        // getters/setters
        ...
    }

Report template
---------------
The [report template](../xls/object_collection_template.xls) for this example uses [Comment markup](../reference/excel_markup.html)  to define the transformation areas.
It looks like this

![Object collection template](../images/object_collection_template.png)

Java code
---------

The Java code uses looks like this 

        List<Employee> employees = generateSampleEmployeeData();
        try(InputStream is = ObjectCollectionDemo.class.getResourceAsStream("object_collection_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/object_collection_output.xls")) {
                Context context = new Context();
                context.putVar("employees", employees);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }

Excel output
------------

Final [report](../xls/object_collection_output.xls) for this example is shown on the following screenshot

![Object collection output](../images/object_collection_output.png)
