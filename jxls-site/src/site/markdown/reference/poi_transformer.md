POI Transformer
===============

[PoiTransformer](http://jxls.sourceforge.net/javadoc/jxls-poi/org/jxls/transform/poi/PoiTransformer.html) 
is an implementation of 
[Transformer](http://jxls.sourceforge.net/javadoc/jxls/org/jxls/transform/Transformer.html) interface 
based [Apache POI](https://poi.apache.org/).

It is a built-in Excel transformer from [jxls-poi](https://bitbucket.org/leonate/jxls/src/master/jxls-poi/) module.

It supports both [POI-HSSF and POI-XSSF/SXSSF workbooks](https://poi.apache.org/components/spreadsheet/) and has multiple constructors
allowing to create a streaming or non-streaming transformer instance from POI workbook or from the template input stream.

Additionally it allows to ignore the source cells row height and/or column width during the cells transformation.
It is achieved with `setIgnoreRowProps(boolean)` and `setIgnoreColumnProps(boolean)` methods.






   

