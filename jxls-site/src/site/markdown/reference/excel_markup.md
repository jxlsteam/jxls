Excel mark-up
=============

Excel markups in Jxls can be split in 3 parts

* Bean properties markup
* Area definition markup
* Command definition markup

Jxls provides *XlsCommentAreaBuilder* class which can read the markup from Excel cell comments.
*XlsCommentAreaBuilder* implements generic *AreaBuilder* interface.

    public interface AreaBuilder {
        List<Area> build();
    }

It is a simple interface with a single method which should return a list of *Area* objects.

So if you want to define your own markup you may create your own implementation of *AreaBuilder* and interpret the input Excel template (or any other input) as desired.

Bean properties markup
----------------------
Jxls uses [Apache JEXL](http://commons.apache.org/proper/commons-jexl/reference) expression language to process

In future releases the expression language engine will be made configurable so that it will be possible to replace JEXL with any other expression engine if required.

The JEXL expression language syntax is explained here [JEXL Syntax](http://commons.apache.org/proper/commons-jexl/reference/syntax.html).

Jxls expects that JEXL expression is placed in `${` `}` in XLS template file.

For example the following cell content `${department.chief.age} years` tells Jxls to evaluate `department.chief.age` using JEXL
assuming the `department` object is available in the *Context*. If for example the expression `department.getChief().getAge()`
evaluates to 35 Jxls will put `35 years` in the cell during XlsArea processing.

XLS Area markup
---------------
Jxls area markup is used to define root XlsArea(s) to be processed by Jxls engine.
*XlsCommentAreaBuilder* supports the following syntax for area definition as an Excel cell command

    jx:area(lastCell="<LAST_CELL>")

Here `<LAST_CELL>` defines the right bottom cell of the rectangular area. The first cell is defined by the cell where the Excel comment is put on.

So assuming we have the next comment `jx:area(lastCell="G12")` in cell `A1` the root area will be read as `A1:G12`.

*XlsCommentAreaBuilder* should be used to read all the areas from the template file.
For example the following code snippet reads all the areas into *xlsAreaList* and then saves the first area into *xlsArea* variable

        AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        Area xlsArea = xlsAreaList.get(0);

In most cases defining of a single root XlsArea should be enough.

Command markup
--------------
*Commands* should be defined inside an *XlsArea*.
*XlsCommentAreaBuilder* accepts the following command notation created as an Excel cell comment

    jx:<command_name>(attr1='val1' attr2='val2' ... attrN='valN' lastCell=<last_cell> areas=["<command_area1>", "<command_area2", ... "<command_areaN>"])

`<command_name>` is a command name pre-registered or manually registered in *XlsCommentAreaBuilder*. Currently the following command names are pre-registered
 
* each
* if
* image

Custom commands may be registered manually with `static void addCommandMapping(String commandName, Class clazz)` method of *XlsCommentAreaBuilder* class.

`attr1`, `attr2`,..., `attrN` are the command specific attributes

For example *If-Command*  has `condition` attribute to set a conditional expression.

`<last_cell>` defines the bottom right cell of the command body area. The top left cell is determined by the cell where the command notation is attached.

`<command_area1>`, `<command_area2>`, ... `<command_areaN>` - XLS areas to be passed to the command as parameter.

For example *If-Command* expects the following areas to be defined

* `ifArea` is a reference to an area to output when the *If-command* condition evaluates to *true*
* `elseArea` is a reference to an area to output when the *If-command* condition evaluates to *false* (optional)

So to define the areas for *If-command* its areas attribute may look like this

    areas=["A8:F8","A13:F13"]

In a single cell comment you can define multiple *Commands*.
For example *Each* and *If* command definitions may  look like this

    jx:each(items="department.staff", var="employee", lastCell="F8")
    jx:if(condition="employee.payment <= 2000", lastCell="F8", areas=["A8:F8","A13:F13"])
