Multi Sheet Markup  demo
=========================

Introduction
------------

This example shows how to use excel markup to output a collection into multiple excel worksheets.
The reference information about multiple sheet processing in JXLS can be found
in [Multiple sheet generation](../reference/multi_sheets.html)

In this example we will operate on a list of *Department* containing a list of *Employee* objects

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

The [report template](../xls/multisheet_markup_template.xls) for this example looks like this

![Multi Sheet template](../images/multi_sheet_markup_template.png)

The multi-sheet related markup is in this instruction

    jx:each(items="departments", var="department", lastCell="G10" multisheet="sheetNames")

Here we indicate that each item of a _departments_ collection should be put on a separate sheet
from a list of sheet names in _sheetNames_ variable in the context.

Java code
---------

        List<Department> departments = EachIfCommandDemo.createDepartments();
        logger.info("Opening input stream");
        try (InputStream is = MultiSheetMarkupDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Context context = PoiTransformer.createInitialContext();
                context.putVar("departments", departments);
                context.putVar("sheetNames", Arrays.asList(
                        departments.get(0).getName(),
                        departments.get(1).getName(),
                        departments.get(2).getName()));
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }

In the code above we associate the _sheetNames_ variable in the context with a list of department names
to have each department worksheet to have the same name as the department name.

Be sure to use the StandardFormulaProcessor for multi sheets. It's the default FormulaProcessor since version 2.4.6.

Excel output
------------

Final [report](../xls/multisheet_markup_output.xls) for this example is shown on the following screenshot

![Multi Sheet output](../images/multi_sheet_markup_output.png)

Each department is generated on a separate worksheet.