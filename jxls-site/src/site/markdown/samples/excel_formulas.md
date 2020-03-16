Excel Formulas Sample
======================

Introduction
------------

This example demonstrates how to use formulas in Jxls.

The example can be found in ObjectCollectionFormulasDemo.java.

In this example we will use the same *Employee* objects as in [Output Object Collection](object_collection.html) guide.

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

The [report template](../xls/formulas_template.xls) for this example uses [Comment markup](../reference/excel_markup.html) to define the transformation areas.
The template is very similar to the template from [Output Object Collection](object_collection.html) example except an additional column `Total Payment` to calculate
total payment for each employee and the `Summary` row to evaluate the sum of all the payments and total payments for all the employees.

Also we updated the area boundary in markup comments to add the new row and new column to the processing area.

![Excel formulas template](../images/formulas_template.png)

Java code
---------

This example uses Jxls POI adapter to generate the report.
The Java code for this example looks like this

        logger.info("Running Object Collection Formulas demo");
        List<Employee> employees = ObjectCollectionDemo.generateSampleEmployeeData();
        try(InputStream is = ObjectCollectionFormulasDemo.class.getResourceAsStream("formulas_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/formulas_output.xls")) {
                Context context = new Context();
                context.putVar("employees", employees);
                JxlsHelper.getInstance().processTemplateAtCell(is, os, context, "Result!A1");
            }
        }

Excel output
------------

Final [report](../xls/formulas_output.xls) for this example is shown on the following screenshot

![Formulas output](../images/formulas_output.png)