Grouping Sample
========================

Introduction
-------------

This sample shows how to use grouping with [Each-Command](../reference/each_command.html) .

The *Employee* class looks like this.

    public class Employee {
        private String name;
        private int age;
        private Double payment;
        private Double bonus;

        // getters/setters
        ...
    }

Report template
---------------
The [report template](../xls/grouping_template.xlsx) for this example uses `groupBy` attribute of [Each-Command](../reference/each_command.html)  
to define the grouping.

    jx:each(items="employees" groupBy="name" groupOrder="asc" lastCell="D6")
    
If you write `groupBy` without `groupOrder` no sorting will be done.
Since the `var` attribute is missing the default group name __group_ is used to refer to the grouped collection items   

![Grouping template](../images/grouping_template.png)

Java code
---------

The Java code is listed below 

        List<Employee> employees = generateSampleEmployeeData();
        try(InputStream is = GroupingDemo.class.getResourceAsStream("grouping_template.xlsx")) {
            try (OutputStream os = new FileOutputStream("target/grouping_output.xlsx")) {
                Context context = new Context();
                context.putVar("employees", employees);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
        
    public static List<Employee> generateSampleEmployeeData() throws ParseException {
        List<Employee> employees = new ArrayList<Employee>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
        employees.add( new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15) );
        employees.add( new Employee("Oleg", dateFormat.parse("1973-Apr-30"), 2300, 0.25) );
        employees.add( new Employee("John", dateFormat.parse("1970-Jul-10"), 3500, 0.10) );
        employees.add( new Employee("Neil", dateFormat.parse("1975-Oct-05"), 2500, 0.00) );
        employees.add( new Employee("Maria", dateFormat.parse("1978-Jan-07"), 1700, 0.15) );
        employees.add( new Employee("John", dateFormat.parse("1969-May-30"), 2800, 0.20) );
        employees.add( new Employee("Oleg", dateFormat.parse("1988-Apr-30"), 1500, 0.15) );
        employees.add( new Employee("Maria", dateFormat.parse("1970-Jul-10"), 3000, 0.10) );
        employees.add( new Employee("John", dateFormat.parse("1973-Apr-30"), 1000, 0.05) );
        return employees;
    }
        

Excel output
------------

Final [report](../xls/grouping_output.xlsx) for this example is shown on the following screenshot

![Grouping output](../images/grouping_output.png)
