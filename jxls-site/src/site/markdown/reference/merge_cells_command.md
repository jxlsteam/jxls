*MergeCells-Command*
===================

Contribution by lnk. See issue [#136](https://bitbucket.org/leonate/jxls/issues/136/i-would-like-to-add-this-merge-command).

Command class
-------------

`MergeCellsCommand`


Excel comment syntax
--------------------

<pre>jx:mergeCells(<br/>
lastCell="Merge cell ranges"<br/>
[, cols="Number of columns combined"]<br/>
[, rows="Number of rows combined"]<br/>
[, minCols="Minimum number of columns to merge"]<br/>
[, minRows="Minimum number of rows to merge"]<br/>
)</pre>

Note: This command can only be used on cells that have not been merged. An exception will occur if the scope of
the merged cell exists for the merged cell.


Transformer Support Note
------------------------

Please note that *MergeCells-Command* is currently supported only in POI transformer so you have to use POI if you are using *MergeCells-Command*.
