Parameterized Excel Formulas Example
====================================

Introduction
-------------

This example shows how to use parameterized formulas in your template.
For more information about Jxls formulas check reference [Excel Formulas](../reference/formulas.html)

Sample data
-----------

In this example we will use the same `Employee` objects as in [Output Object Collection](object_collection.html) guide.

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

The [report template](../xls/param_formulas_template.xls) for this example looks like this

![Parameterized Formulas template](../images/param_formulas_template.png)


Java code
---------

In this example we will use Jxls POI adapter to generate the report.
The Java code is the same as in [Excel Formulas example](../samples/excel_formulas.html) except that we are putting `bonus` variable
into our context.

        List<Employee> employees = generateSampleEmployeeData();
        try(InputStream is = ParameterizedFormulasDemo.class.getResourceAsStream("param_formulas_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/param_formulas_output.xls")) {
                Context context = new Context();
                context.putVar("employees", employees);
                context.putVar("bonus", 0.1);
                JxlsHelper.getInstance().processTemplateAtCell(is, os, context, "Result!A1");
            }
        }

Excel output
------------

Final [report](../xls/param_formulas_output.xls) for this example is shown on the following screenshot

![Formulas output](../images/param_formulas_output.png)