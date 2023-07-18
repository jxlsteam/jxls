Jxls 2.13.0 is released!
============================

* [#178 Java 17](https://github.com/jxlsteam/jxls/issues/178) *- Thanks to mosidev*
* [#209 Use of groupBy and select](https://github.com/jxlsteam/jxls/issues/209)
  <br/>**Breaking change:** If jx:each groupBy  _and_  select are specified
  the filtering will now be done  _before_  the grouping.
  **".item" must then be removed from the select expression.**
  We think the old processing order makes no sense and is used by nobody.
  If you really need the old processing order give us an information in
  [#209](https://github.com/jxlsteam/jxls/issues/209) and set jx:each parameter oldSelectBehavior="true".
* [#235 POI 5.2.2](https://github.com/jxlsteam/jxls/issues/235) *- Thanks to Sivakumar-finoli*
  <br/>We advise against using POI 5.2.3. POI 5.2.2 can be used. An Apache POI bug ticket [66687](https://bz.apache.org/bugzilla/show_bug.cgi?id=66687) has been created. Bug came from [65562](https://bz.apache.org/bugzilla/show_bug.cgi?id=65562).
  <br/>We've also created [66679](https://bz.apache.org/bugzilla/show_bug.cgi?id=66679) because of potential
  problems with cell type General.
* [#237 made getSheetsNameOfMultiSheetTemplate() public](https://github.com/jxlsteam/jxls/issues/237)
* [#244 JxlsNationalLanguageSupport problems](https://github.com/jxlsteam/jxls/issues/244)


###### The latest component versions

* org.jxls:jxls:2.13.0

* org.jxls:jxls-poi:2.13.0

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
