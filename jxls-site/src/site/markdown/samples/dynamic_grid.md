Dynamic Grid Sample
======================

Introduction
------------
The example shows how to output a grid with dynamic number of columns/rows using [Grid Command](../reference/grid_command.html).

The example can be found in GridCommandDemo.java.

In this example we will use the following simple Java bean 

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

The [report template](../xls/grid_template.xls) for this example looks like this

![Grid Command template](../images/grid_template.png)
 
It defines the Grid-Command in cell `A3` using the following cell comment
 
    jx:grid(lastCell="A4" headers="headers" data="data" areas=[A3:A3, A4:A4] formatCells="BigDecimal:C1,Date:D1")
 
Here the header area is `A3:A3` and the data area is `A4:A4`.

And the body of the command in the template file contains the following two cells

    ${header}
    ${cell}
    
Java code
---------
The example uses Jxls POI adapter to generate the report.
The Java code looks like this

    private static void executeGridObjectListDemo(List<Employee> employees) throws IOException {
        try(InputStream is = GridCommandDemo.class.getResourceAsStream("grid_template.xls")) {
            try(OutputStream os = new FileOutputStream("target/grid_output2.xls")) {
                Context context = new Context();
                context.putVar("headers", Arrays.asList("Name", "Birthday", "Payment"));
                context.putVar("data", employees);
                JxlsHelper.getInstance().processGridTemplateAtCell(is, os, context, "name,birthDate,payment", "Sheet2!A1");
            }
        }
    }

The list of headers is set like this
 
    context.putVar("headers", Arrays.asList("Name", "Birthday", "Payment"));
    
And the data for grid body is a list of *Employee* objects.

        context.putVar("data", employees);

And we pass *Employee* object property names to be used for the grid to *processGridTemplateAtCell* method
    
        JxlsHelper.getInstance().processGridTemplateAtCell(is, os, context, "name,payment,birthDate", "Sheet2!A1");

Excel output
------------

Final for this example is shown on the following screenshots 

[Sheet1](../xls/grid_output1.xls)

![Grid output - Sheet1](../images/grid_output1.png)

[Sheet2](../xls/grid_output2.xls)

![Grid output - Sheet2](../images/grid_output2.png)