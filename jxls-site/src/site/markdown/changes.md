Version History
==============

v2.12.0
--------
* [#147 Row height bugfix], contribution by [jools-uk](https://github.com/jools-uk)


v2.11.0
--------
The release contains an upgrade of JEXL library and a fix for XXE vulnerability

* [#131 JEXL 3.2](https://github.com/jxlsteam/jxls/issues/131)
* [#143 Vulnerability alert](https://github.com/jxlsteam/jxls/issues/143)


v2.10.0
--------
New features and fixed bugs

* [#78 Add support for JSR310 types (aka Java Time API)](https://github.com/jxlsteam/jxls/issues/78), contribution by [wagnerluis1982](https://github.com/wagnerluis1982)
* [#79 Multi-line and comment support for SQL queries in jx:each](https://github.com/jxlsteam/jxls/issues/79), contribution by [alexlust](https://github.com/alexlust)
* [#85 Distinguish between unknown key and run var](https://github.com/jxlsteam/jxls/issues/85)
* [#87 Formula recalculation](https://github.com/jxlsteam/jxls/issues/87), contribution by [Turbocube644](https://github.com/Turbocube644)
* [#93 Backup/restore varIndex in jx:each](https://github.com/jxlsteam/jxls/issues/93)
* [#108 bug in 2.10.0 RC1: Don't call getRunVar with null](https://github.com/jxlsteam/jxls/issues/108)
* [#111 bug in 2.10.0 RC2: LiteralsExtractor buggy](https://github.com/jxlsteam/jxls/issues/111)

other resolved issues:

* [#75 Documented varIndex argument](https://github.com/jxlsteam/jxls/issues/75), contribution by [sapradhan](https://github.com/sapradhan)
* [#90 Escaping single quotes in query is possible](https://github.com/jxlsteam/jxls/issues/90)

v2.9.0
-----------
With this version we migrated from BitBucket to Github.

A list of resolved issues

* [#1: Javadoc](https://github.com/jxlsteam/jxls/issues/1)
* [#2: moved jxls-site into main repo](https://github.com/jxlsteam/jxls/issues/2)
* [#8: Gradle](https://github.com/jxlsteam/jxls/issues/8)
* [#13: Code coverage](https://github.com/jxlsteam/jxls/issues/13)
* [#15: Java 8](https://github.com/jxlsteam/jxls/issues/15)
* [#19: B prefix for BitBucket issue related testcase classes](https://github.com/jxlsteam/jxls/issues/19)
* [#25: read-only mode when open Excel files in tests](https://github.com/jxlsteam/jxls/issues/25)
* [#27: changes from 2.8.1](https://github.com/jxlsteam/jxls/issues/27), see below and [BitBucket issue 210](https://bitbucket.org/leonate/jxls/issues/210/sum-when-more-than-1-sheet-doesnt-work-on)
* [#50: POI 4.1.2, commons-beanutils 1.9.4, slf4j](https://github.com/jxlsteam/jxls/issues/50)
* [#53: CLOB fix](https://github.com/jxlsteam/jxls/issues/53)
* [#64: don't read iterable ahead of time (jx:each bugfix)](https://github.com/jxlsteam/jxls/issues/64)
* [#68: workaround for POI bug: unset value after changing a formula (streaming mode)](https://github.com/jxlsteam/jxls/issues/68)

v2.8.1
-----------
The release contains a critical bug-fix for a specific case of formulas processing. 
See [issue#210 SUM with empty lists when more than 1 sheet doesn't work](https://bitbucket.org/leonate/jxls/issues/210/sum-when-more-than-1-sheet-doesnt-work-on) for more detail.


v2.8.0
-----------

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

v2.7.2
--------
This release fixes [issue#188 Referencing other sheet in JXLS-processed cell formula replaces formula with "=0"](https://bitbucket.org/leonate/jxls/issues/188/referencing-other-sheet-in-jxls-processed) .

Also it reverts back the behaviour from previous Jxls releases (2.6.0 and earlier) where the POI formulas evaluation was not triggered by default.
See corresponding section on [Excel Formulas](reference/formulas.html) page for more information on how to enable POI formulas evaluation if you need it.

v2.7.1
-----------
This is a bugfix release fixing the following issue
* [issue#184 jx:if within jx:each may cause problem](https://bitbucket.org/leonate/jxls/issues/184/jx-if-within-jx-each-may-cause-problem)

v2.7.0
-----------
Besides many bug-fixes (see below) this release has the following changes

* Added POI 4.1.0 compatibility
* Introduced varIndex variable to get iteration index in jx:each
* Removed Jexcel module

* [issue#72 Group sum](https://bitbucket.org/leonate/jxls/issues/72) (contribution by Marcus Warm)
* [issue#97 'jx:each' command not processing Division formula having cell ref outside each](https://bitbucket.org/leonate/jxls/issues/97/jx-each-command-not-processing-division)
* [issue#108 setDeleteTemplateSheet(true) doesn't work](https://bitbucket.org/leonate/jxls/issues/108/setdeletetemplatesheet-true-doesnt-work)
* [issue#127 Problem with formulas referencing cells in another worksheet containing special characters in name](https://bitbucket.org/leonate/jxls/issues/127/problem-with-formulas-referencing-cells-in)
* [issue#132 delete multi sheet template](https://bitbucket.org/leonate/jxls/issues/132/delete-multi-sheet-template)
* [issue#149 Pre-evaluating (excel) formulas during template processing](https://bitbucket.org/leonate/jxls/issues/149/pre-evaluating-excel-formulas-during)
* [issue#157 jx:each select attribute limited expression evaluation context](https://bitbucket.org/leonate/jxls/issues/157/jx-each-select-attribute-limited)
* [issue#158 Row can be null in SelectSheetsForStreamingPoiTransformer (NPE)](https://bitbucket.org/leonate/jxls/issues/158) (contribution by Marcus Warm)
* [issue#159 Insert image issues](https://bitbucket.org/leonate/jxls/issues/159/insert-image-and-text-underline-issues) (contribution by ZhengJin Fang)
* [issue#160 jx:each with direction=RIGHT with SXSSF Transformer rewrites static cells](https://bitbucket.org/leonate/jxls/issues/160/jx-each-with-direction-right-with-sxssf)
* [issue#161 Multisheet names longer than 31 letters are not gracefully handled](https://bitbucket.org/leonate/jxls/issues/161/multisheet-names-longer-than-31-letters)
* [issue#162 Bug in parameterized Excel Formulas](https://bitbucket.org/leonate/jxls/issues/162/bug-in-parameterized-excel-formulas)
* [issue#164 Remove Jexcel](https://bitbucket.org/leonate/jxls/issues/164/remove-jexcel)
* [issue#166 Excel AVERAGE function in template not working properly](https://bitbucket.org/leonate/jxls/issues/166/excel-average-function-in-template-not)
* [issue#168 POI 4.1.0 will raise errors on org.jxls.area.XlsArea#transformTopStaticArea](https://bitbucket.org/leonate/jxls/issues/168/poi-410-will-raise-errors-on)
* [issue#169 Trim comment before checking if it's a JxlsxParams comment](https://bitbucket.org/leonate/jxls/issues/169/trim-comment-before-checking-if-its-a)
* [issue#173 how can I get index in jx:each](https://bitbucket.org/leonate/jxls/issues/173/how-can-i-get-index-in-jx-each)
* [issue#174 NullPointer thrown if cell has empty comment](https://bitbucket.org/leonate/jxls/issues/174/nullpointer-thrown-if-cell-has-empty)
* [issue#180 Nested sums](https://bitbucket.org/leonate/jxls/issues/180/nested-sums)


v2.6.0
-----------

* 2.6.x is the last version with jxls-jexcel support. Please use jxls-poi.
* [issue#136 Added MergeCellsCommand](https://bitbucket.org/leonate/jxls/issues/136/i-would-like-to-add-this-merge-command) (contribution by lnk)
* [issue#138 Upgrade JEXL to v3.1](https://bitbucket.org/leonate/jxls/issues/138/upgrade-jexl)
  <br/>You must change the package from jexl2 to jexl3.
  Now adding custom functions works like this:
  `evaluator.setJexlEngine(new JexlBuilder().namespaces(functionMap).create());`
* [issue#139 Upgrade 3rd-party dependencies](https://bitbucket.org/leonate/jxls/issues/139/upgrade-third-party-dependencies)
* [issue#142 Increase test coverage](https://bitbucket.org/leonate/jxls/issues/142/increase-test-coverage)
* [issue#143 Removed AreaCommand.clearCells](https://bitbucket.org/leonate/jxls/issues/143/areacommandclearcells)
* [issue#144 If/else-Each-If/else testcase](https://bitbucket.org/leonate/jxls/issues/144/if-else-each-if-else-testcase)
* [issue#137 Moved jxls-poi sources into jxls repo](https://bitbucket.org/leonate/jxls/issues/137/monorepo-jxls-and-jxls-poi)
* [issue#140 Moved jxls-demo sources into jxls repo](https://bitbucket.org/leonate/jxls/issues/140/mono-repo-jxls-demo-into-jxls)
* [issue#141 Improved code quality](https://bitbucket.org/leonate/jxls/issues/141/improve-code-quality-part-1) (contribution by Leonid and Marcus)

v2.5.1
-----------

Added a default implementation for the Transformer.adjustTableSize method to make jxls-2.5.x compatible with the previous releases of jxls-poi

v2.5.0
-----------

This release features [Table support](https://bitbucket.org/leonate/jxls/issues/88) and an option to use an expression for multi-sheet generation.

Thanks to *Marcus Warm* and *zangloo* for the contributions!

jxls-2.5.0

* fixed [issue#88 Table support](https://bitbucket.org/leonate/jxls/issues/88)
* Possibility to use an expression in `multisheet` property for sheet name (see [Multiple sheets output](reference/multi_sheets.html) for more info)
* fixed [issue#117 Grouping null values](https://bitbucket.org/leonate/jxls/issues/117)
* fixed [issue#129 Usage of Multi Sheet Markup with empty sheet names list causes IndexOutOfBoundsException](https://bitbucket.org/leonate/jxls/issues/129)
* fixed [issue#134 Optimize Context.getVar(String)](https://bitbucket.org/leonate/jxls/issues/134/optimization-in-contextgetvar-string)

jxls-poi-1.1.0

* fixed issue#88 Table support
* fixed [issue#104 Extending PoiTransformer is hard](https://bitbucket.org/leonate/jxls/issues/104)
* fixed [issue#14 POI 4.0.1 support](https://bitbucket.org/leonate/jxls-poi/issues/14/poi-401)

jxls-jexcel-1.0.8

* issue#88 No table support for jxls-jexcel

v2.4.7
-----------
jxls-2.4.7
 
* fixed [issue#125 Allow passing null image to ImageCommand](https://bitbucket.org/leonate/jxls/issues/125/image-command-customization)
* fixed [issue#124 Optimizing XlsArea:findCommandsForVerticalShift recursive method](https://bitbucket.org/leonate/jxls/issues/124/optimizing-xlsarea)
* fixed [issue#122 Wrong cell replacement](https://bitbucket.org/leonate/jxls/issues/122/wrong-cell-replacement)

jxls-poi-1.0.16

* support for Apache POI 4.0.0

jxls-reader-2.0.5

* support for Apache POI 4.0.0
* reverted the fix for "XLSBlockReader startRow/endRow not work" merged earlier in [PR#5](https://bitbucket.org/leonate/jxls-reader/pull-requests/5/fixed-for-xlsblockreader-startrow-endrow/diff) 

v2.4.6
-----------
* [issue#111 StandardFormulaProcessor is now the default formula processor for XlsArea and JxlsHelper](https://bitbucket.org/leonate/jxls/issues/111)
* [issue#116 Formula handling issues (formula external to any jx:area)](https://bitbucket.org/leonate/jxls/issues/116/formula-handling-issues-formula-external)

v2.4.5
-----------
The following issues have been fixed

* [issue#100 Cell format not being correctly shifted with an empty list](https://bitbucket.org/leonate/jxls/issues/100/cell-format-not-being-correctly-shifted)
* [issue#103 Cell format not being correctly shifted when having JXLS command and empty list](https://bitbucket.org/leonate/jxls/issues/103/cell-format-not-being-correctly-shifted)
* [issue#105 Issue with 'big double values' (like 1.3E22) being parsed as cell references (like E22)](https://bitbucket.org/leonate/jxls/issues/105/issue-with-big-double-values-like-13e22)
* [issue#106 Issue with row height and jxls2 root not in A1](https://bitbucket.org/leonate/jxls/issues/106/issue-with-row-height-and-jxls2-root-not)
* [issue#107 Issue with formatting of parts of text](https://bitbucket.org/leonate/jxls/issues/107/issue-with-formatting-of-parts-of-text)

v2.4.4
-----------
* [issue#26 FastFormulaProcessor has a buggy replacement strategy](https://bitbucket.org/leonate/jxls/issues/26/fastformulaprocessor-has-a-buggy)

v2.4.3
-------
* [issue#92 Make JxlsHelper extensible](https://bitbucket.org/leonate/jxls/issues/92/make-jxlshelper-extensible) 
* [issue#84 Allow using of Iterables in Each command](https://bitbucket.org/leonate/jxls/issues/84/useful-collection-iterable)
* [issue#8 Jxls is not compatible with POI 3.17](https://bitbucket.org/leonate/jxls-poi/issues/8/poitransformer-does-not-compatible-with)
* [issue#9 Avoid using deprecated POI API](https://bitbucket.org/leonate/jxls-poi/issues/9/resolve-deprecated-poi-api)

v2.4.2
------
This a bug-fix release fixing

* [issue#85: Wrong sum in multisheet](https://bitbucket.org/leonate/jxls/issues/85/wrong-sum-in-multisheet) 

v2.4.1
------
This release features several bug-fixes and a few improvements

* [issue#74: Use Resultsets on JXLS library](https://bitbucket.org/leonate/jxls/issues/74) 
* [issue#77: JXLS Losing rows when using if-command inside each-command](https://bitbucket.org/leonate/jxls/issues/77) 
* [issue#80: Cannot export more than 256 columns into .xlsx](https://bitbucket.org/leonate/jxls/issues/80) 
* [issue#78: ConvertUtils is not thread safe](https://bitbucket.org/leonate/jxls/issues/78) 
* [issue#81: ignore collection parse error in each-command](https://bitbucket.org/leonate/jxls/issues/81) 
* [issue#82: restore the original varName object value in each-command](https://bitbucket.org/leonate/jxls/issues/82) 
* [Using ServiceFactory SPI to load ExpressionEvaluator implementation](https://bitbucket.org/leonate/jxls/pull-requests/35) 


v2.4.0
-------
This release features

* [UpdateCellCommand](reference/updatecell_command.html) for cell processing customization
* Grouping support in [Each-Command](reference/each_command.html) 
 
The following bugs have been fixed 

* [issue#58: Using sxssfTransformer formulas weren't generated properly](https://bitbucket.org/leonate/jxls/issues/58/using-sxssftransformer-formulas-werent)
* [issue#42: Formula processing mechanism](https://bitbucket.org/leonate/jxls/issues/42/formula-processing-mechanism)
* [issue#59: SUM with more than 255 parameters](https://bitbucket.org/leonate/jxls/issues/59/sum-with-more-than-255-parameters)
* [issue#53: NullPointerException is thrown when IfCommand condition is false](https://bitbucket.org/leonate/jxls/issues/53/nullpointerexception-is-thrown-when)
* [issue#65: NullPointerException is thrown when excel template have multiple MergedRegions](https://bitbucket.org/leonate/jxls/issues/65/nullpointerexception-is-thrown-when-excel)
* [issue#54: jxls 2.3.0 memory leak when using on Tomcat](https://bitbucket.org/leonate/jxls/issues/54/jxls-230-memory-leak-when-using-on-tomcat)
* [issue#64: "areas" regex does not handle whitespace](https://bitbucket.org/leonate/jxls/issues/64/areas-regex-does-not-handle-whitespace)
* [issue#33: can not processFormulas while use createSxssfTransformer](https://bitbucket.org/leonate/jxls/issues/33/can-not-processformulas-while-use)

v2.3.0
------
Now you can use a markup in Excel template to output a collection into multiple sheets.
See [Multiple sheets output](reference/multi_sheets.html) section and [Multi sheet markup example](samples/multi_sheet_markup_demo.html)

Also the following issues have been fixed

* [issue#35 Each command problem with row contain if command](https://bitbucket.org/leonate/jxls/issues/35/each-command-problem-with-row-contain-if)
* [issue#23 Each-command with direction "RIGHT" doesn't shift outer columns](https://bitbucket.org/leonate/jxls/issues/23/each-command-with-direction-right-doesnt)
* [issue#32 Deletion of sheet fails](https://bitbucket.org/leonate/jxls/issues/32/deletion-of-sheet-fails)
* [issue#45 Incorrect row height is set in Each-command](https://bitbucket.org/leonate/jxls/issues/45/incorrect-row-height-is-set-in-each)
* [issue#25 Row heights reset between parameterized cells](https://bitbucket.org/leonate/jxls/issues/25/row-heights-reset-between-parameterized)
* [jxls-poi issue#1 PoiTransformer MAX_COLUMN_TO_READ_COMMENT should be configurable property](https://bitbucket.org/leonate/jxls-poi/issues/1/poitransformer-max_column_to_read_comment)
* [jxls-poi issue#2 Printing context to log in case of Exception and huge context pollutes log file](https://bitbucket.org/leonate/jxls-poi/issues/2/printing-context-to-log-in-case-of)
* [jxls-poi issue#3 Display gridlines property is not correctly copied](https://bitbucket.org/leonate/jxls-poi/issues/3/display-gridlines-property-is-not)
* [jxls-poi issue#4 Provide Shared String Table Option for SXSSF Transformer](https://bitbucket.org/leonate/jxls-poi/issues/4/provide-shared-string-table-option-for)
* [jxls-jexcel issue#1 CellType.DATE is not written to a template](https://bitbucket.org/leonate/jxls-jexcel/issues/1/celltypedate-is-not-written-to-a-template)

There are also the following api changes

* Method `createInitialContext()` was removed from  _Transformer_ interface

    The static method of the same name in _PoiTransformer_  and _JexcelTransformer_ can now be used for the same purpose (e.g. `PoiTransformer.createInitialContext()`)

* `transform` method in _Transformer_ interface is now taking an additional parameter `boolean updateRowHeight` indicating if a row height needs to be updated

* A new method `void updateRowHeight(String srcSheetName, int srcRowNum, String targetSheetName, int targetRowNum)` is added to
_Transformer_ interface to copy row height from a source row to a target row.

v2.2.9
------
The following issues were fixed

* Fixed [issue#34](https://bitbucket.org/leonate/jxls/issues/34/each-command-problem-with-the-merge-cells) Each-command problem with the merge cells
* Fixed [issue#36](https://bitbucket.org/leonate/jxls/issues/36/huge-performance-issue-with-debug-logging) Huge performance issue with debug logging
* Fixed [issue#37](https://bitbucket.org/leonate/jxls/issues/37/dont-obtrude-a-logging-framework-conflicts) Don't obtrude a logging framework

v2.2.8
-------
This release features some bug-fixes and improvements

* An option to set a default formula value (see [Excel formulas](reference/formulas.html) for more details)
* Class api improvements by adding method chaining to *JxlsHelper* and *CellRef* classes (contributed by Karri-Pekka Laakso)
* Fixed [issue#24](https://bitbucket.org/leonate/jxls/issues/24/outofmemoryerror-on-big-data) 
by adding *isFormulaProcessingRequired* attribute to *Context.Config*
* Fixed [issue#29](https://bitbucket.org/leonate/jxls/issues/29/jxls-each-right-issue-missing-datas)
* Fixed [issue#27](https://bitbucket.org/leonate/jxls/issues/27/deletion-of-sheet-fails)
* Fixed [issue#30](https://bitbucket.org/leonate/jxls/issues/30/sum-formula-gives-n-a-when-no-row-in-the)


v2.2.7
-----------
* Added an option to use SQL queries in Excel Template. See [SQL usage in template](reference/sql_in_template.html)
* Added an option to specify *formula strategy* for any given formula allowing to adjust formula calculation  
(you can find a demo of this functionality in *FormulaCopyDemo* example in jxls-demo project)

v2.2.6
-----------
Fixed an error when processing an area with multiple nested commands [issue#21](https://bitbucket.org/leonate/jxls/issues/21/fix-problem-with-incorrect-processing-of)


v2.2.5
-----------
Improvements for [Image-command](reference/image_command.html)

* *imgBean* attribute is now renamed to *src* and supports any expression resulting in image byte array *byte[]*


v2.2.4
-----------
* [Each-command](reference/each_command.html) now supports setting of *direction* attribute in Excel template with a text value *DOWN* or *RIGHT*
* Minor refactoring: replacing *StringBuffer* with *StringBuilder*

v2.2.3
------
This is a bug-fix release.
The following issues were fixed

* [issue #9 No obvious behavior using Each-Command](https://bitbucket.org/leonate/jxls/issues/9/no-obvious-behavior-using-each-command)
* [issue #11 Problem with complex formulas](https://bitbucket.org/leonate/jxls/issues/11/problem-with-complex-formulas)


v2.2.2
------
This is a bug-fix release.
The following issues were fixed

* [issue #8 Bug at using nested formulas](https://bitbucket.org/leonate/jxls/issues/8/bug-at-using-nested-formulas)
* [issue #10 Final width of a grid is not calculated properly](https://bitbucket.org/leonate/jxls/issues/10/final-width-of-a-grid-is-not-calculated)

v2.2.1
------
* Better support for [SXSSF](https://poi.apache.org/spreadsheet/how-to.html#sxssf) templates
* *JxlsHelper* class to simplify the library usage. [Getting Started](getting_started.html) guide and some of the examples were updated to use *JxlsHelper*.  

v2.2.0
------
###New features
* Possibility to replace JEXL with another expression language engine (see [expression language](reference/expression_language.html)
* Possibility to replace the default `${}` expression notation with the one you like (e.g. `[[]]`)

###Bugs fixed
* [Bug when using right direction each command](https://bitbucket.org/leonate/jxls/issues/7/bug-when-using-right-direction-each)

v2.1.1
-------
This release introduces [SimpleExporter](reference/simple_exporter.html) that allows you 
to generate an excel with a single line of code and without necessity to provide a template file.

v2.1.0
------
This release introduces new [Grid Command](../reference/grid_command.html) and contains the following bug-fixes

* [NPE in IfCommand create Excel Markup](https://bitbucket.org/leonate/jxls/issue/2/npe-in-ifcommand-create-excel-markup)
* [Free rows when 2 adjacent collections](https://bitbucket.org/leonate/jxls/issue/4/free-rows-when-2-adjacent-collections)


v2.0.0
------
First public release for Jxls-2.

Jxls-2 is a full rewrite from scratch of the original [Jxls 1.x](http://jxls.sf.net/1.x) library to make it much more flexible and improve the performance.

Jxls-2 completely de-couples itself from the underlying Java-to-Excel transformation engine so that it is possible to switch low-level Java-to-Excel library without any code changes.

Excel template markup can now be redefined in a way you like.

Jxls-2 introduces a concept of [Command](reference/command.html) to replace previously used tags (like *jx:forEach*, *jx:if*).
[Command](reference/command.html) is not tied to any particular Excel mark-up. In fact it can be defined completely in Java or by using built-in Excel-comment or XML-based mark-ups.

Custom post- and pre- processing code can now also be easily injected into Excel transformation.

Please note that Jxls-2 is not backwards compatible with Jxls 1.x.

If you wish to upgrade your existing Jxls 1.x reports for Jxls-2 you will have to rewrite the code and change the notation in XLS template.