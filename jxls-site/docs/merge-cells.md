# Merge cells

```
jx:mergeCells(cols="" rows="" minCols="" minRows="" lastCell="C2")
```

`cols`: Number of columns combined

`rows`: Number of rows combined

`minCols`: Minimum number of columns to merge

`minRows`: Minimum number of rows to merge

`lastCell`: Merge cell ranges

This command can only be used on cells that have not been merged. An exception will occur if the scope of
the merged cell exists for the merged cell.

jx:mergeCells is part of jxls-poi and is only available if you use JxlsPoiTemplateFillerBuilder or add the MergeCellsCommand using withCommand().

This command is a community contribution.