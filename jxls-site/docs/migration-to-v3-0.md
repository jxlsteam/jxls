# Migration from version 2.x to 3.0 guide

## Why version 3.0?

We wanted to switch from Java 8 to **Java 17**. A **Logback** update was necessary for this. This proved difficult as the
Logback XML feature (Joran) has changed massively. We therefore decided to remove the XML feature and make a major release 3.
We believe that no one has used the XML feature. (If this isn't the case, the community is welcome to recreate the XML feature for
v3 - with a better XML technology.)

We ended up removing Logback entirely. We used version 3 for further extensive revisions, while the core remained unchanged.

**SLF4J** has also been removed and replaced with the **JxlsLogger interface**.

The **new builder API** replaces **JxlsHelper**. All options can now be set in the builder with a Fluent API.
This now makes it possible to create reports with completely different options in parallel.
The TransformerFactory and createTransformer() methods have also been removed.

**Context** is now no longer just seen as a Map holder, but should in the future contain all the information necessary for creating reports.
Context is now an **internal** interface. The data is passed to the JxlsTemplateFiller as Map&lt;String, Object>; i.e. without Context.
Sometimes classes need Context access. There is the PublicContext interface and the builder method needsPublicContext() for this.
TransformationConfig was renamed to ExpressionEvaluatorContext and moved from the Transformer to the Context.
This means that expression evaluations are now done via the Context.

Another design goal was the possibility of exchanging classes.

A new strategy is that it should be possible to add new commands without having to change the Transformer.

The **documentation** has been completely rewritten and our website has a new design.

In addition to the "note lastCell" command syntax from version 2, we have a new alternative "note marker" syntax in the pipeline. See MarkerAreaBuilder.

**Sorry** for the trouble the updating will take. But Jxls is a hobby project and the hard transition was less effort.
**The templates do not need to be changed.**
You can stay with 2.14.0 *for a while*. If there are important bug fixes, we will also make them there if necessary.

## Overview

**Removals**

- Java 8 support (official support for Java 17)
- Logback, creating commands with XML
- ServiceLoader mechanism removed for less trouble with Docker.
- Groovy testcases
- PoiContext

**Replacements**

- JxlsLogger replaces SLF4J logging (but see our SLF4J page for an adapter class)
- new builder API replaces JxlsHelper
- JxlsTransformerFactory replaces TransformerFactory (and other static createTransformer() methods)
- Map&lt;String, Object> replaces Context
- grid_template.xls was changed to grid_template.xlsx
- JdbcHelper was renamed to DatabaseAccess.

**New Features**

- ASC_ignoreCase, DESC_ignoreCase options new for jx:each/orderBy and jx:each/groupOrder
- Activate streaming for a sheet using `sheetStreaming="true"` in a note (see JxlsStreaming.AUTO_DETECT)
- new builder options e.g. pre write actions

## What have to be changed in your code?

**TL;DR:** use Map&lt;String, Object> instead of Context and use JxlsPoiTemplateFillerBuilder instead of JxlsHelper.

I had to switch to Jxls 3 myself and here are these tips from the **experience** I had:

- Before switching to version 3, you should centralize the Jxls calls if you have not already done so. We also recommend writing a unit test for (almost) every report. We do it like this.
- Instead of using `JxlsHelper.getInstance()` use `JxlsPoiTemplateFillerBuilder.newInstance()`. (See [builder options](builder.html) documentation)
- Instead of Context Use Map&lt;String, Object> (called "data").
- You could use MapDelegator&lt;String, Object> as base class if you need special Map behavior. Just overwrite one or some methods.
- Call Map.put() instead of Context.putVar(). If you have many calls like this you can create your own HashMap implementation with a putVar() method.
- clear() method of JexlExpressionEvaluator is now static. You don't need to navigate via Transformer and TransformationConfig.
- If you add a custom functions object to the data map you can now use `builder.needsPublicContext(customFunctions)` and implement NeedsPublicContext interface on your custom functions class to retrieve the PublicContext. Evaluations can be done using the PublicContext.
- If you used `new JexlExpressionEvaluator(expression).evaluate(context.toMap());` you can now use `context.evaluate(expression)`
- If you used `JxlsHelper.getInstance().createExpressionEvaluator(null).evaluate(condition, context.toMap())` to evaluate a condition which returns a boolean you can now use: `new ExpressionEvaluatorFactoryJexlImpl().createExpressionEvaluator(null).isConditionTrue(condition, data)`
- If you want to execute code before `transformer.write()` you can call `builder.withPreWriteAction((transformer, publicContext) -> ...)`
- If you want to get the Excel report as byte[] you could write this: `builder.buildAndFill(data)` and get a byte[]
- If you want a strict non-silent expression evaluator do this: `builder.withExpressionEvaluatorFactory(new ExpressionEvaluatorFactoryJexlImpl(true));` and also do this: `builder.withLogger(new PoiExceptionThrower())` for getting exceptions if something went wrong. Old Jxls 2 code was like this for that: `transformer.getTransformationConfig().setExpressionEvaluator(new JexlExpressionEvaluator(false, true)); transformer.setExceptionHandler(new PoiExceptionThrower());`
- The central expression evaluator is now also used by jx:each/select. If you have the central expression evaluator set to nonsilent-strict, Jxls 3 may now throw EvaluationExceptions. The expressions then need to be corrected.
- If you want to use SLF4J then see our builder/logger/SLF4J page for a ready to use adapter class.
- If you called `context.getConfig().setIsFormulaProcessingRequired(false);` you must now call `builder.withUpdateCellDataArea(false)`
- The old option evaluateFormulas is now called recalculateFormulasBeforeSaving and can be set with the builder.
- The old option fullFormulaRecalculationOnOpening is now called recalculateFormulasOnOpening and can be set with the builder.
- Instead of creating SelectSheetsForStreamingPoiTransformer you can just use `builder.withStreaming(JxlsStreaming.AUTO_DETECT)`.
- PoiUtil (accessed in templates using "util.") and SafeSheetNameBuilder are no longer part of the Context. You now have to add them yourself.
- Util.toByteArray() has been moved to ImageCommand. There's also a useful `copy(InputStream, OutputStream)` method in ImageCommand. You possibly have to change your templates.
- SafeSheetNameBuilder.createSafeSheetName() has a new 3rd argument: JxlsLogger logger
