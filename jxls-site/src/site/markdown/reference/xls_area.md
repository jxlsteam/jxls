XLS Area
========

Introduction
------------

XLS Area is a major concept in JxlsPlus.
Basically it represents a rectangular area in an Excel file which needs to be transformed.
Each XLS Area may have a list of transformation Commands associated with it and a set of nested child areas.
Each child area is also an XLS Area with its own set of Commands and nested areas.
A top-level XLS Area is an area which does not have a parent area (it is not nested into any other XLS Area).

Constructing XLS Area
---------------------

There are 3 ways to build XLS Area

* Using Excel markup

* Using XML configuration

* Using Java API

Let's describe each of these methods in detail

### Excel markup to build XLS Area

You can use a special markup in your Excel template to construct XLS area.
The markup should be placed into an Excel comment for the first cell of the area.
The markup looks like this

    jx:area(lastCell = "<AREA_LAST_CELL>")

where `<AREA_LAST_CELL>` is the last cell of the defined area.

This markup defines a top-level area starting from the cell with the markup comment and ending  in the `<AREA_LAST_CELL>` .

To see an example let's look at the template from the [Output Object Collection](../samples/object_collection.html) sample

![Object collection template](../images/object_collection_template.png)

An area is defined in a comment for `A1` cell as

    jx:area(lastCell="D4")

So here we have an area covering `A1:D4` cell range.

 To parse the markup and create `XlsArea` object we should use *XlsCommentAreaBuilder* class as shown below

    // getting input stream for our report template file from classpath
    InputStream is = ObjectCollectionDemo.class.getResourceAsStream("object_collection_template.xls");
    // creating POI Workbook
    Workbook workbook = WorkbookFactory.create(is);
    // creating JxlsPlus transformer for the workbook
    PoiTransformer transformer = PoiTransformer.createTransformer(workbook);
    // creating XlsCommentAreaBuilder instance
    AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
    // using area builder to construct a list of processing areas
    List<Area> xlsAreaList = areaBuilder.build();
    // getting the main area from the list
    Area xlsArea = xlsAreaList.get(0);

The following 2 code lines do all the main job on area building

    AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
    List<Area> xlsAreaList = areaBuilder.build();

First you are constructing *AreaBuilder* instance by instantiating *XlsCommentAreaBuilder*.
And the second step is to invoke `areaBuilder.build()` method to construct a list of *Area* objects from the template.

After you get a list of top-level areas you can then use them for Excel transformation.

### XML Configuration to build XLS Area

If you prefer to define XLS Area with XML markup here is how you can do this.

First you have to create an XML configuration defining your area.

As an example let's consider a simple XML configuration from the [Object Collection output with XML Builder](../samples/object_collection_xmlbuilder.html)

    <xls>
        <area ref="Template!A1:D4">
            <each items="employees" var="employee" ref="Template!A4:D4">
                <area ref="Template!A4:D4"/>
            </each>
        </area>
    </xls>

The root element is `xls`. Then you can list a number of `area` elements defining each of the top-level areas.

 Here we have a single top-level area `A1:D4` at `Template` sheet.

    <area ref="Template!A1:D4">

Inside the area we define associated commands using a special element for a specific command. In this case we are defining
`each command` with `each` xml element. The area associated with the `each command` is indicated  with `ref` attribute.

    <each items="employees" var="employee" ref="Template!A4:D4">

Inside the `each command` we have a nested area parameter as

    <area ref="Template!A4:D4"/>

### Java API to build XLS Area

To create an XLS Area with Java API you may use one of the `XlsArea` class constructors. The following constructors are available

     public XlsArea(AreaRef areaRef, Transformer transformer);

     public XlsArea(String areaRef, Transformer transformer);

     public XlsArea(CellRef startCell, CellRef endCell, Transformer transformer);

     public XlsArea(CellRef startCellRef, Size size, List<CommandData> commandDataList, Transformer transformer);

     public XlsArea(CellRef startCellRef, Size size);

     public XlsArea(CellRef startCellRef, Size size, Transformer transformer);

To build a top level area you have to provide a `Transformer` instance so that the area can use it for transformation.

And you have to define the area cells using cell range as a string or alternatively by creating `CellRef` cell reference object and set the area `Size`.

Here is a snippet of code to construct a set of nested template XLS areas with commands

    // create Transformer instance
    // ...
    // Create a top level area
    XlsArea xlsArea = new XlsArea("Template!A1:G15", transformer);
    // Create 'department' are
    XlsArea departmentArea = new XlsArea("Template!A2:G13", transformer);
    // create 'EachCommand' to iterate through departments
    EachCommand departmentEachCommand = new EachCommand("department", "departments", departmentArea);
    // create an area for employee 'each' command
    XlsArea employeeArea = new XlsArea("Template!A9:F9", transformer);
    // create an area for 'if' command
    XlsArea ifArea = new XlsArea("Template!A18:F18", transformer);
    // create 'if' command with the specified areas
    IfCommand ifCommand = new IfCommand("employee.payment <= 2000",
            ifArea,
            new XlsArea("Template!A9:F9", transformer));
    // adding 'if' command instance to employee area
    employeeArea.addCommand(new AreaRef("Template!A9:F9"), ifCommand);
    // create employee 'each' command and add it to department area
    Command employeeEachCommand = new EachCommand( "employee", "department.staff", employeeArea);
    departmentArea.addCommand(new AreaRef("Template!A9:F9"), employeeEachCommand);
    // add department 'each' command to top-level area
    xlsArea.addCommand(new AreaRef("Template!A2:F12"), departmentEachCommand);


