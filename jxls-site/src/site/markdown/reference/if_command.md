*If-Command*
============

Introduction
------------

*If-Command* is a conditional command to output an area depending on a condition specified in the `test` attribute of the command.

Command Attributes
------------------

The *If-Command*  has the following attributes

* `condition` is a conditional expression to test

* `ifArea` is a reference to an area to output when this command condition evaluates to *true*

* `elseArea` is a reference to an area to output when this command condition evaluates to *false*

* `lastCell` is a common attribute for any command pointing to the last cell of the command area

`ifArea` and `condition` attributes are mandatory.

Command building
----------------

As with any Jxls command you can use Java API or Excel markup or XML configuration to define the *If-Command*

### Java API usage

Below is an example of creating *If-Command* found in org.jxls.examples

    // ...
    // creating 'if' and 'else' areas
    XlsArea ifArea = new XlsArea("Template!A18:F18", transformer);
    XlsArea elseArea = new XlsArea("Template!A9:F9", transformer);
    // creating 'if' command
    IfCommand ifCommand = new IfCommand("employee.payment <= 2000", ifArea, elseArea);

### Excel markup

To create *If Command* with Excel markup you should use the following syntax in a comment for a starting cell of the command body area

    jx:if(condition="employee.payment <= 2000", lastCell="F9", areas=["A9:F9","A18:F18"])

Here `lastCell` attribute defines the last cell of the *If-Command* area.

### XML markup

To create *If-Command* with XML configuration you use the following markup

    <area ref="Template!A9:F9">
        <if condition="employee.payment &lt;= 2000" ref="Template!A9:F9">
            <area ref="Template!A18:F18"/>
            <area ref="Template!A9:F9"/>
        </if>
    </area>

Here `ref` attribute defines the area to be associated with the *If-Command*.

