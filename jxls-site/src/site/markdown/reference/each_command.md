Each-Command
==============

Introduction
------------

*Each-Command* is used to iterate through a collection and clone the command XLS area.
It is an analogue of Java *for* operator.

Command Attributes
------------------

*Each-Command* has the following attributes

* `var` is a name of the variable in Jxls context to put each new collection item when iterating

* `items` is a name of a context variable containing the collection (Iterable<?>) or array to iterate

* `varIndex` is name of variable in Jxls context that holds current iteration index, zero based 

* `area` is a reference to XLS Area used as `each command` body

* `direction` is a value of `Direction` enumeration which may have values `DOWN` or `RIGHT` to indicate how to repeat the command body - by rows or by columns. The default value is `DOWN`.

* `select` is an expression selector to filter out collection items during the iteration

* `groupBy` is a property to do the grouping

* `groupOrder` indicates ordering for groups ('desc' or 'asc')

* `orderBy` contains the names separated with comma and each with an optional postfix " ASC" (default) or " DESC" for the sort order

* `cellRefGenerator` is a custom strategy for target cell references creation

* `multisheet` is a name of a context variable containing a list of sheet names to output the collection

* `lastCell` is a common attribute for any command pointing to the last cell of the command area

The `var` and `items` attributes are mandatory while others can be skipped.

To find more information about using `cellRefGenerator` and `multisheet` attributes check [Multiple sheets section](multi_sheets.html).

Command building
----------------

As with any Jxls command you can use Java API or Excel markup or XML configuration to define the *Each-command*

# Java API usage

Below is an example of creating *Each-Command* found in package org.jxls.examples.

    // creating a transformer and departments collection
        ...
    // creating department XlsArea
        XlsArea departmentArea = new XlsArea("Template!A2:G13", transformer);
    // creating Each Command to iterate departments collection and attach to it "departmentArea"
        EachCommand departmentEachCommand = new EachCommand("department", "departments", departmentArea);

# Excel markup usage

To create *Each Command* with Excel markup you should use a special syntax in the comment for a starting cell of the command body area

    jx:each(items="employees" var="employee" lastCell="D4")

So we use `jx:each` with the command attributes in parentheses separated with a space. `lastCell` attribute defines the last cell of the command XlsArea body.

# XML markup usage

To create *Each Command* with XML configuration you use the following markup

    <each items="employees" var="employee" ref="Template!A4:D4">
        <area ref="Template!A4:D4"/>
    </each>


Here `ref` attribute defines the area to be associated with the *Each-Command*. And the inner area defines the body of the *Each-Command*.
Usually they are the same.

Repeat direction
----------------

By default *Each-Command* `direction` attribute is set to `DOWN` which means that the command body will be cloned down over Excel rows.

If you need to clone the area by columns you should set the `direction` attribute to `RIGHT` value.

With Java API you do this like

    //... creating EachCommand to iterate departments
    // setting its direction to RIGHT
    departmentEachCommand.setDirection(EachCommand.Direction.RIGHT);
    
Grouping the data
------------------
*Each-Command* supports grouping via its `groupBy` property. The `groupOrder` property sets the ordering and can be `desc` or `asc`.
If you write `groupBy` without `groupOrder` no sorting will be done.

In the excel markup it can look like this

    jx:each(items="employees" var="myGroup" groupBy="name" groupOrder="asc" lastCell="D6")
    
In this example each group can be referred using _myGroup_ variable which will be accessible in the context.
    
The current group item can be referred using  `myGroup.item`. So to refer to employee name use

    ${myGroup.item.name}
    
All the items in the group are accessible via `items` property of the group e.g.    

    jx:each(items="myGroup.items" var="employee" lastCell="D6")
    
You can also skip the `var` attribute altogether and in this case the default group variable name will be __group_.    
    
For an example look at the [Grouping Example](../samples/grouping_example.html)     
