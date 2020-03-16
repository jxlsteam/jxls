Multiple Sheet CellRefGenerator  demo
======================================

Introduction
------------

This example shows how to use output a collection into multiple excel worksheets.
More information can be found in [Multiple sheet generation](../reference/multi_sheets.html)

In this example we will use *Department* and *Employee* objects

    public class Department {
        private String name;
        private Employee chief;
        private List<Employee> staff = new ArrayList<Employee>();
        private String link;

        // getters/setters
        ...
    }

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

The [report template](../xls/multi_sheet_template.xls) for this example looks like this

![Multi Sheet template](../images/multi_sheet_template.png)


Java code
---------

In this example we will use Jxls POI transformer to generate the report. If necessary you can easily modify it to use Jexcel transformer.

        List<Department> departments = EachIfCommandDemo.createDepartments();
        logger.info("Opening input stream");
        try(InputStream is = EachIfCommandDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                System.out.println("Creating area");
                XlsArea xlsArea = new XlsArea("Template!A1:G15", transformer);
                XlsArea departmentArea = new XlsArea("Template!A2:G12", transformer);
                EachCommand departmentEachCommand = new EachCommand("department", "departments", departmentArea, new SimpleCellRefGenerator());
                XlsArea employeeArea = new XlsArea("Template!A9:F9", transformer);
                XlsArea ifArea = new XlsArea("Template!A18:F18", transformer);
                IfCommand ifCommand = new IfCommand("employee.payment <= 2000",
                        ifArea,
                        new XlsArea("Template!A9:F9", transformer));
                employeeArea.addCommand(new AreaRef("Template!A9:F9"), ifCommand);
                Command employeeEachCommand = new EachCommand("employee", "department.staff", employeeArea);
                departmentArea.addCommand(new AreaRef("Template!A9:F9"), employeeEachCommand);
                xlsArea.addCommand(new AreaRef("Template!A2:F12"), departmentEachCommand);
                Context context = new Context();
                context.putVar("departments", departments);
                logger.info("Applying at cell Sheet!A1");
                xlsArea.applyAt(new CellRef("Sheet!A1"), context);
                xlsArea.processFormulas();
                logger.info("Complete");
                transformer.write();
                logger.info("written to file");
            }
        }

Excel output
------------

Final [report](../xls/multi_sheet_output.xls) for this example is shown on the following screenshot

![Multi Sheet output](../images/multi_sheet_output.png)

Each department is generated on a separate worksheet.