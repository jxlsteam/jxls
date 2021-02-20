Jxls 2.10.0-rc1 is available!
==========================

The release candidate build for Jxls 2.10 is available for download.

New features:

* [#78 Add support for JSR310 types (aka Java Time API)](https://github.com/jxlsteam/jxls/issues/78), contribution by [wagnerluis1982](https://github.com/wagnerluis1982)
* [#79 Multi-line and comment support for SQL queries in jx:each](https://github.com/jxlsteam/jxls/issues/79), contribution by [alexlust](https://github.com/alexlust)
* [#85 Distinguish between unknown key and run var](https://github.com/jxlsteam/jxls/issues/85)
* [#87 Formula recalculation](https://github.com/jxlsteam/jxls/issues/87), contribution by [Turbocube644](https://github.com/Turbocube644)
* [#93 Backup/restore varIndex in jx:each](https://github.com/jxlsteam/jxls/issues/93)

other resolved issues:

* [#75 Documented varIndex argument](https://github.com/jxlsteam/jxls/issues/75), contribution by [sapradhan](https://github.com/sapradhan)
* [#90 Escaping single quotes in query is possible](https://github.com/jxlsteam/jxls/issues/90)
 

###### The latest component versions

* org.jxls:jxls:2.10.0-rc1

* org.jxls:jxls-poi:2.10.0-rc1

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
