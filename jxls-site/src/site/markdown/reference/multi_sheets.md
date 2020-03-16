# Multiple sheets

You can create sheets at runtime. Just add a `jx:each` command and use the `multisheet` attribute. Each item in jx:each represents a sheet.

## Given names

    jx:each(items="departments", var="dep", multisheet="sheetnames", lastCell="D4")

The `items` define how many sheets are created at runtime. The names come from a `String List` that was put to the `Context`
with the var name specified by the `multisheet` attribute (here: "sheetnames").

This was implemented by the `SheetNameGenerator` class used by `EachCommand`.

## Dynamic names

    jx:each(items="departments", var="dep", multisheet="dep.name", lastCell="D4")

The `items` define how many sheets are created at runtime. The sheet name comes from the expression defined in the `multisheet` attribute.

This was implemented by the `DynamicSheetNameGenerator` class used by `EachCommand`.

## Use PoiContext to get safe and unique sheet names

`PoiContext` contains a `PoiSafeSheetNameBuilder` instance. `PoiSafeSheetNameBuilder` ensures valid and unique sheet names.
If there are less sheet names (defined in the `multisheet` var) than defined sheets (defined in the `items` attribute)
the `PoiSafeSheetNameBuilder` will create sheet names. An INFO message will be printed to the log if a sheet name was modified.

If you just use `Context` there will be no `SafeSheetNameBuilder` by default. If there are not enough sheet names or a sheet name
is not valid or not unique an ERROR message will be printed to the log. The sheet won't be created!

## Cookbook

Here are some recipes how to change the sheet name.
Always extend `PoiSafeSheetNameBuilder` and put the instance with the name `SafeSheetNameBuilder.CONTEXT_VAR_NAME`
to the `Context`.

### Change the first serial number

Override the method `getFirstSerialNumber`.

### Change the sheet name layout

Override the method `addSerialNumber`.

Example: If there are two sheets with the name "data", the 2nd one will get the name "data(1)". With the following code you can change
it to "data-2".

        context.putVar(SafeSheetNameBuilder.CONTEXT_VAR_NAME, new PoiSafeSheetNameBuilder() {
            @Override
            protected int getFirstSerialNumber() {
                return 2;
            }
            
            @Override
            protected String addSerialNumber(String text, int serialNumber) {
                return text + "-" + serialNumber;
            }
        });

### Add a serial number at begin of each sheet name

Override the method `createSafeSheetName`, modify the `givenSheetName` and call the super method.

        context.putVar(SafeSheetNameBuilder.CONTEXT_VAR_NAME, new PoiSafeSheetNameBuilder() {
            @Override
            public String createSafeSheetName(String givenSheetName, int index) {
                return super.createSafeSheetName((index + 1) + ". " + givenSheetName, index);
            }
        });

### Delete template sheet

TBD

## Use your own CellRefGenerator

If you directly use the `[EachCommand](each_command.html)` object instead of `jx:each` you can set your
own `CellRefGenerator` instance and have full control over generating sheet names.

Let's take a look at an example for multi-sheet output found in [Multi sheet demo](../samples/multi_sheet_demo.html) example.

    // create transformer and defining command areas
    ...
    // creating each command providing custom cell reference generator instance
    EachCommand departmentEachCommand = new EachCommand("department", "departments", departmentArea, new SimpleCellRefGenerator());
    // define other commands and areas
    ...
    // adding command to an area, setting up the bean context and transforming the template
    xlsArea.addCommand(new AreaRef("Template!A2:F12"), departmentEachCommand);
    Context context = new Context();
    context.putVar("departments", departments);
    logger.info("Applying at cell Sheet!A1");
    xlsArea.applyAt(new CellRef("Sheet!A1"), context);
    // finishing the processing
    ...

As can be seen we provided an instance of `SimpleCellRefGenerator` to `EachCommand`.
The code for `SimpleCellRefGenerator` is very simple. We just have to implement a single method to return a cell
reference where to start output a Department data for each iteration.


    public class SimpleCellRefGenerator implements CellRefGenerator {
        public CellRef generateCellRef(int index, Context context) {
            return new CellRef("sheet" + index + "!B2");
        }
    }

Our implementation just returns a new sheet reference for each iteration so the first department will go into
sheet0!B2, second into sheet1!B2 and so on.
