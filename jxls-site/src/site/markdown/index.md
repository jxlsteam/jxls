Jxls v2.8.0 is released!
========================

The release brings many bug-fixes and some improvements, in particular

* [National language support](reference/nls.html)
* Excel conditional formatting support for `jx:each` command [issue #110](https://bitbucket.org/leonate/jxls/issues/110/conditional-formatting-rules-are-not)
* JxlsTester for writing template based testcases. See [Jxls source page](source_code.html) for more information.
* [Order-by for jx:each command](https://bitbucket.org/leonate/jxls/issues/193/add-orderby-to-jx-each-command)

A list of resolved issues

* [issue#90 Allow Excel formulas to work with jx:grid](https://bitbucket.org/leonate/jxls/issues/90/allow-excel-formulas-to-work-with-jx-grid) contribution by Matthias Lung
* [issue#133 Group by nested property](https://bitbucket.org/leonate/jxls/issues/133/group-by-nested-property)
* [issue#155 Table and PivotTable support does not work with dynamic column names](https://bitbucket.org/leonate/jxls/issues/155/table-and-pivottable-support-does-not-work) = National language support
* [issue#183 Util.transformToIterableObject should say whether a var is null](https://bitbucket.org/leonate/jxls/issues/183/utiltransformtoiterableobject-should-say) 
* [issue#185 Better error message for corrupt Excel file](https://bitbucket.org/leonate/jxls/issues/185/better-error-message-for-corrupt-excel)
* [issue#186 Standardized testcases](https://bitbucket.org/leonate/jxls/issues/186/standardized-testcases)
* [issue#193 Add orderBy to jx:each command](https://bitbucket.org/leonate/jxls/issues/193/add-orderby-to-jx-each-command)
* [issue#195 Add logback-classic for tests](https://bitbucket.org/leonate/jxls/issues/195/add-logback-classic-for-tests)
* [issue#196 groupBy and groupOrder support for XML markup](https://bitbucket.org/leonate/jxls/issues/196/groupby-and-grouporder-support-for-xml)
* [issue#197 Jointed cell references don't work anymore with several empty collections](https://bitbucket.org/leonate/jxls/issues/197/jxls-with-version-272-jointed-cell)
* [issue#198 Array support for EachCommand](https://bitbucket.org/leonate/jxls/issues/198/array-support-for-eachcommand)
* [issue#200 Treat null in EachCommand items as empty collection](https://bitbucket.org/leonate/jxls/issues/200/treat-null-in-eachcommand-items-as-empty) 
* [issue#201 Fix broken links](https://bitbucket.org/leonate/jxls/issues/201/fix-broken-links)
* [issue#204 Order of conditional formatting rules is not preserved on copying](https://bitbucket.org/leonate/jxls/issues/204/order-of-conditional-formatting-rules-is)
* [issue#206 Each with RIGHT next to each other does not work](https://bitbucket.org/leonate/jxls/issues/206)
* [issue#207 IndexOutOfBoundsException in StandardFormulaProcessor](https://bitbucket.org/leonate/jxls/issues/207/indexoutofboundsexception-in)


###### Compatibility notes and future plans
* Signature of createTransformer() has been changed. The throws has been removed. Now the unchecked CannotOpenWorkbookException will be thrown.
* jx:each and jx:grid now treat null lists as empty lists (see [issue#200](https://bitbucket.org/leonate/jxls/issues/200/treat-null-in-eachcommand-items-as-empty) )
* POI 4.0 requires Java 8 and so it is the recommended version to use. JXLS code base will be migrated to Java 8 syntax in the upcoming versions.
* The major change in 2.9.0 will be the move to Github. Our homepage and publishing to Maven-central will stay the same.
 

###### The latest component versions

* org.jxls:jxls:2.8.0

* org.jxls:jxls-poi:2.8.0

Introduction
------------
Jxls is a small Java library to make generation of Excel reports easy.
Jxls uses a special markup in Excel templates to define output formatting and data layout.

Excel generation is required in many Java applications that have some kind of reporting functionality.

Java has a few libraries for creating Excel files e.g. [Apache POI](https://poi.apache.org/).

Those libraries are great but quite low-level as they require a developer to write a lot of Java code even to create a simple Excel file.

Usually one has to manually set each cell formatting and data for the spreadsheet.
Depending on the complexity of the report layout and data formatting the Java code can become quite complex and difficult to debug and maintain.

In addition not all Excel features are supported and can be manipulated with libraries API (e.g. limited support for macros, graphs etc).
The suggested workaround for unsupported features is to create an object manually in an Excel template  and fill in the template with data after that. Jxls takes this approach to a higher level. 

When working with Jxls one just needs to define the required report formatting and data layout in an Excel template file and then run Jxls engine
 to fill in the template with data. A developer needs to write just a little bit of Java code to trigger Jxls engine processing of the template.

Features
--------
* XML and binary Excel format output (depends on underlying low-level Java-to-Excel implementation)
* Java collections output by rows and by columns
* Conditional output
* [Expression language](reference/expression_language.html) in report definition markup 
* [Multiple sheets output](reference/multi_sheets.html)
* Native Excel formulas
* Parameterized formulas
* Grouping support
* Merged cells support
* Area listeners to adjust excel generation
* Excel comments mark-up for command definition
* XML mark-up for command definition
* Custom Command definition
* Streaming for fast output and less memory consumption
* Streaming for selected sheets (SelectSheetsForStreamingPoiTransformer)
* Table support
