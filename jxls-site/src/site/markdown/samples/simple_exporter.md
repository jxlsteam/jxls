SimpleExporter Demo
===================

Introduction
------------

The example shows how to use  [SimpleExporter](../reference/simple_exporter.html).

The Java object for this example looks like this 

    public class Employee {
        private String name;
        private Date birthDate;
        private BigDecimal payment;
        private BigDecimal bonus;
        // constructors and getters/setters
        .....
    }

Report template
---------------
By default [SimpleExporter](../reference/simple_exporter.html) uses a [built-in template](../xls/builtin_template.xls).

But in this example we also use a custom template to demonstrate how to customize the built-in template.
Our custom template for *SimpleExporter* is [here](../xls/simple_export_template.xlsx) and looks like this

![Simple Export template](../images/simple_export_template.png)
 
Java code
---------
You should use POI transformer to use *SimpleExporter*.
The Java code for this example looks like this

        try(OutputStream os1 = new FileOutputStream("target/simple_export_output1.xls")) {
            List<Employee> employees = generateSampleEmployeeData();
            List<String> headers = Arrays.asList("Name", "Birthday", "Payment");
            SimpleExporter exporter = new SimpleExporter();
            exporter.gridExport(headers, employees, "name, birthDate, payment", os1);

            // now let's show how to register custom template
            try (InputStream is = SimpleExporterDemo.class.getResourceAsStream(template)) {
                try (OutputStream os2 = new FileOutputStream("target/simple_export_output2.xlsx")) {
                    exporter.registerGridTemplate(is);
                    headers = Arrays.asList("Name", "Payment", "Birth Date");
                    exporter.gridExport(headers, employees, "name,payment, birthDate,", os2);
                }
            }
        }

There are two invocations of *gridExport* method.

In the first case the built-in template is used 

        exporter.gridExport(headers, employees, "name, birthDate, payment", os1);
        
In the second case we first register our custom template with 

        InputStream is = SimpleExporterDemo.class.getResourceAsStream(template);
        exporter.registerGridTemplate(is);

And then perform the same *gridExport* method invocation but with different properties order     
   
        exporter.gridExport(headers, employees, "name,payment, birthDate,", os2);
   
 
Excel output
------------

The final Excel generated from the built-in template is [here](../xls/simple_export_output1.xls) and looks like this

![Built-in template usage output](../images/simple_export_output1.png)

The Excel for the custom registered template is [here](../xls/simple_export_output2.xlsx) and looks like this

![Custom template usage output](../images/simple_export_output2.png)