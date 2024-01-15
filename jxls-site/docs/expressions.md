# Expressions

Expressions in cells are inside `${...}`. Everything inside is evaluated by an Expression Evaluator.
Some command attributes also contain expressions - but there without `${...}`.

Jxls supports these expression languages:

- Apache JEXL (default, built-in)
- JSR 223

Other expression languages could be integrated by implementing ExpressionEvaluatorFactory and ExpressionEvaluator.
And using the builder method withExpressionEvaluatorFactory().

## JEXL {{jexl_version}}

See [homepage](https://commons.apache.org/proper/commons-jexl/index.html) and especially the
[JEXL syntax reference](https://commons.apache.org/proper/commons-jexl/reference/syntax.html).

### Some examples

Access a property: `obj.propertName`

Add to values: `obj.number1+obj.number2`

condition: `obj.number1 > obj.number2 && obj.number3 < 100`

If-else: `obj.condition ? "true value" : "else value"`

Is empty() and size(): `empty(list) ? "There are no items." : "There are " + size(list) + " item(s)."`

Prevent null: `obj.propertyName??""`

## JSR 223

Expression evaluators based on [JSR 223](https://www.jcp.org/en/jsr/detail?id=223) *"Scripting for the Java Platform"* are also supported,
e.g. the [Spring Expression Language (SpEL)](https://docs.spring.io/spring-framework/reference/core/expressions.html).
