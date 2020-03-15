Excel Formulas
==============

Jxls-2 supports standard Excel formulas as well as parameterized formulas defined with a special syntax and 
allowing you to use the context parameters in a formula. 

## Triggering Jxls formula processing

If you use [JxlsHelper](http://jxls.sourceforge.net/javadoc/jxls/org/jxls/util/JxlsHelper.html)  to process the template 
the formulas will be processed by default so no code changes are needed

    JxlsHelper.getInstance().processTemplateAtCell(is, os, context, "Result!A1");
    
Formula processing can be disabled by setting `processFormulas` flag to `false` 

    JxlsHelper.getInstance().setProcessFormulas(false);    

### Formulas evaluation

If you want to enable formula evaluation (see https://poi.apache.org/components/spreadsheet/eval-devguide.html) you should do it explicitly
with `setEvaluateFormula(true)`

    JxlsHelper.getInstance().setEvaluateFormulas(true).processTemplate(is, os, context);    
    
If you create a Transformer instance yourself and want to enable formula evaluation you should call `setEvaluateFormulas(true)`
on the transformer instance itself

        Transformer transformer = PoiTransformer.createTransformer(in, out);
        transformer.setEvaluateFormulas(true);
        JxlsHelper.getInstance().processTemplate(context, transformer);


## Formulas Template example 


To see how Jxls formulas look in a template file let's take a template from [Excel formulas example](../samples/excel_formulas.html).

You can download it from  [here](../xls/formulas_template.xls).

The screenshot of the template is below

![Excel formulas template](../images/formulas_template.png)

As you can see there are three formula cells in the template `E4`, `C5`, `E5` .

Cell `E4` is itself a part of `jx:each` command area and contains a formula `=C4*(1+D4)`.

Both `C4` and `D4` cells are part of the same `jx:each` command area.

After jx:each transformation the area `A4:E4` will be expanded into multiple rows.

The original formula `=C4*(1+D4)` will be modified accordingly for each new row so that we will get formulas like `=C5*(1+D5)`, `=C6*(1+D6)` and so on.

The summation formula in cell `C5` `=SUM(C4)` references cell `C4` from inside `jx:each` command area.

It means that after transformation and formulas processing the reference to cell `C4` will be replaced with a range like `SUM(C4:C8)`.

Same goes for cell `E5` with formula `=SUM(E4)`.

The final excel output is  [here](../xls/formulas_output.xls) and also can be seen on the following screenshot

![Formulas output](../images/formulas_output.png)
     
    
## Use XlsArea to trigger formulas processing

If you do not use [JxlsHelper](http://jxls.sourceforge.net/javadoc/jxls/org/jxls/util/JxlsHelper.html) and want to process formulas
in your template you should explicitly call  `processFormulas()` method of XlsArea instance after the main transformation is done.

When executing `processFormulas()`  Jxls engine will process and render all the template formulas
updating them as necessary to count for possible cells shifting and collections expanding.

The Java code may look like this

    Area xlsArea;
    Context context;
    // construct XLS Area and set it into xlsArea var
    // ...
    // fill in context var with data
    // ...
    // apply XLS Area at A1 cell of 'Result' sheet
    xlsArea.applyAt(new CellRef("Result!A1"), context);
    // process area formulas
    xlsArea.processFormulas();
    // save excel output
    // ...

The line `xlsArea.processFormulas()` does all the formulas processing job. 


## Formula Processor

By default Jxls uses *StandardFormulaProcessor* to process formulas in the template when *processFormulas()* method is invoked.

For trivial templates you could also use the *FastFormulaProcessor* which performs 10 times faster. You must call

     xlsArea.setFormulaProcessor(new FastFormulaProcessor()); 

before *processFormulas()* is invoked.

Example that uses JxlsHelper:

     JxlsHelper.getInstance().setUseFastFormulaProcessor(true).processTemplate(...)


## Parameterized Formulas

Parameterized formula allows you to use context variables in the formula.

To set a parameterized formula you have to enclose it into `$[` and `]` symbols and each formula variable must be enclosed in `$\{` and `}` symbols.
For example `$[SUM(E4) * ${bonus}]` . Here we use 'bonus' context variable in the formula.
During the `processFormulas()` Jxls will substitute all the variables with values from the context.

To see this in action please take a look at [Parameterized formulas example](../samples/param_formulas.html)


## Default Formula value

If the cells participating in a formula calculation are removed during processing then the formula value can become corrupted or undefined.
To avoid this situation starting from v.2.2.8 Jxls sets such formula to *=0*.
To use custom default value for such formulas use *jx:params* comment to set *defaultValue* property.
For example

    jx:params(defaultValue="1")

This sets the default formula value to 1.


## Jointed cell references

If your formula refers to cells from different areas which should be combined into a single range or cell sequence you
may use so called *jointed* cell reference notation combined with the parameterized formulas e.g.

    $[SUM(U_(F8,F13))]

Note `U_()` notation which will tells Jxls to use both cells `F8` and `F13` and 
combine the target cells into a single cell sequence or range if possible.

To see a working example of this notation take a look at the template comment_markup_demo.xls used by XlsCommentBuilderDemo.java.

## Cell reference tracking

While performing an area transformation Jxls keeps track of all the processed cells so that it knows what are the target cells for each particular source cell.
If you do not have or do not need to process the formulas then it makes sense to disable this functionality to save some memory.
This can be done by setting the following configuration parameter into the context config e.g.

        Context context = new PoiContext();
        context.getConfig().setIsFormulaProcessingRequired(false);