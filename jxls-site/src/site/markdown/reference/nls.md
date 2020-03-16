# National language support

Since JXLS 2.8.0 you can use resource bundle keys in your Excel templates to realize multilingualism.
Just write

    R{counterpart.name}

in your template cell. For English language it will become "Counterpart" and for German language it will be translated into "Kontrahent".
This feature is **not activated by default** and works only for **.xlsx** files. You have to pre-process your template using this code:

    final java.util.Properties resourceBundle = ...
    JxlsNationalLanguageSupport nls = new JxlsNationalLanguageSupport() {
       @Override
       protected String translate(String name, String fallback) {
           return resourceBundle.getProperty(name, fallback);
       }
    };
    File temp = nls.process("template.xlsx"));

After that call your JXLS code (usually `JxlsHelper`) to fill the template file (temp). Finally, the temp file has to be deleted by you.
You must implement `translate()`. so it's totally in your hand to get the value for the given key (`name`). If you have no value you
can return the `fallback` value. It's the value specified as the default value (see below) or it's just the same value as `name`.

## Default value

You can write the language for your default language into the template by using this syntax:

    R{counterpart.name=Counterpart}

## Pivot tables

This NLS feature is especially important for Excel pivot tables. There you can *not* use a NLS solution like `${RB.msg("counterpart.name")}`
because it will break the link between the pivot tables and your data (e.g. a named table).

## Change syntax

It's possible to change the syntax calling the JxlsNationalLanguageSupport setters.

    JxlsNationalLanguageSupport nls = ...
    nls.setStart("{{");
    nls.setEnd("}}");
    nls.setDefaultValueDelimiter(": ");


## See also

- PivotTableTest.java
- JxlsNationalLanguageSupport.java
- BitBucket JXLS issue 155
