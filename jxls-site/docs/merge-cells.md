# Merge cells <!-- ** -->

```
jx:mergeCells(cols="" rows="" minCols="" minRows="" condition="" lastCell="C2")
```

`cols`: Number of columns combined

`rows`: Number of rows combined

`minCols`: Minimum number of columns to merge

`minRows`: Minimum number of rows to merge

`condition`: Optional JEXL expression. If set and evaluates to false, the area is processed but cells are not merged. If omitted or true, merge behaves as before.

`lastCell`: Merge cell ranges

This command can only be used on cells that have not been merged. An exception will occur if the scope of
the merged cell exists for the merged cell.

jx:mergeCells is part of jxls-poi and is only available if you use JxlsPoiTemplateFillerBuilder or add the MergeCellsCommand using withCommand().

This command is a community contribution.

Example (merge department column only when a group has more than one row):

```
jx:each(items="employees" groupBy="buGroup" groupOrder="asc" lastCell="B2")
jx:mergeCells(cols="1" rows="_group.items.size()" condition="_group.items.size() > 1" lastCell="A2")
```
