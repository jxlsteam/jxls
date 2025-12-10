# Create reports <!-- ** -->

An Excel report is generated from

- Jxls [options](builder.html),
- the Jxls [template](template.html) and
- data.

Since version 3, the options are put together with the new JxlsTemplateFillerBuilder with a fluent API.

The template is then added using a withTemplate() method.

Call build() and fill(data, outputFile) for creating an Excel report.

```
Map<String, Object> dataMap = new HashMap<>(); // and fill dataMap...
JxlsTemplateFiller filler = JxlsPoiTemplateFillerBuilder.newInstance()
    // builder options...
    .withTemplate("filename.xlsx")
    .build();
filler.fill(dataMap, new JxlsOutputFile(new File("report.xlsx")));
```

Or shorter:

```
Map<String, Object> dataMap = new HashMap<>(); // and fill dataMap...
JxlsPoiTemplateFillerBuilder.newInstance()
    // builder options...
    .withTemplate("filename.xlsx")
    .buildAndFill(dataMap, new File("report.xlsx"));
```

## Data <!-- ** -->

The data is passed to Jxls as a single `Map<String, Object>` object.
A map value can be a scalar, an Iterable, an array or an *object*.
Scalar: String, Double, boolean and so on.
*Object* can be a POJO class or a `Map<String, Object>`.
Jxls must be able to add entries to the map and remove entries from the map.

In version 2 Jxls used a class called Context that held a `Map<String, Object>`.
In version 3 we have simplified it to `Map<String, Object>` and call this object 'data map'.

## Common options <!-- ** -->

Common options are [streaming](streaming.html) and throwing exceptions. And it's recommended to use [permissions](builder.html#factory-options).

```
JxlsPoiTemplateFillerBuilder.newInstance()
    .withExceptionThrower()
    .withExpressionEvaluatorFactory(new ExpressionEvaluatorFactoryJexlImpl(
        false, true, JxlsJexlPermissions.RESTRICTED))
    .withStreaming(JxlsStreaming.AUTO_DETECT)
    .withTemplate("filename.xlsx")
    .buildAndFill(dataMap, new JxlsOutputFile(new File("report.xlsx")));
```

If you want to use SLF4J for logging you must replace `withExceptionThrower()` by `withLogger(new Slf4jJxlsLogger())`
and copy the class from the [SLF4J](slf4j.html) page.


## Real world production scenario <!-- ** -->

In real world production scenario you want to have a stable codebase. You can't deliver every day. It should be possible just to change
a template file. So it's a philosophy to have as much as possible inside the template files.
