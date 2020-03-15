Object Collection Output with XML Builder
=========================================

Introduction
------------

This sample shows how to output an object collection using XML configuration file to define XLS transformation areas.

We will use a list of the following `Employee` objects to demonstrate how to output an object collection to Excel with JxlsPlus.

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

The [report template](../xls/object_collection_xmlbuilder_template.xls) is very simple and does not include any markup to define transformation areas.
It looks like this

![Object collection XML builder template](../images/object_collection_xmlbuilder_template.png)

XML Configuration
-----------------

The XML configuration file for this example looks like this

    <xls>
        <area ref="Template!A1:D4">
            <each items="employees" var="employee" ref="Template!A4:D4">
                <area ref="Template!A4:D4"/>
            </each>
        </area>
    </xls>

Java code
---------

In this example we will use jXLS POI adapter to generate the report.
The Java code is listed below

        List<Employee> employees = generateSampleEmployeeData();
        try(InputStream is = ObjectCollectionXMLBuilderDemo.class.getResourceAsStream("object_collection_xmlbuilder_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/object_collection_xmlbuilder_output.xls")) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                try (InputStream configInputStream = ObjectCollectionXMLBuilderDemo.class.getResourceAsStream("object_collection_xmlbuilder.xml")) {
                    AreaBuilder areaBuilder = new XmlAreaBuilder(configInputStream, transformer);
                    List<Area> xlsAreaList = areaBuilder.build();
                    Area xlsArea = xlsAreaList.get(0);
                    Context context = new Context();
                    context.putVar("employees", employees);
                    xlsArea.applyAt(new CellRef("Result!A1"), context);
                    transformer.write();
                }
            }
        }

Excel output
------------

Final [report](../xls/object_collection_xmlbuilder_output.xls) with all the required cells highlighted is shown on the following screenshot

![Object collection XML builder output](../images/object_collection_xmlbuilder_output.png)
