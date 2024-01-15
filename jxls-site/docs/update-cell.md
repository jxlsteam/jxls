# Update cell

Update cell allows you to customize the processing for a particular cell using this interface:

```
public interface CellDataUpdater {
    void updateCellData(CellData cellData, CellRef targetCell, Context context);
}
```

```
jx:updateCell(updater="myUpdater" lastCell="E2")
```


`updater`: property name for getting CellDataUpdater instance

`lastCell`: area end

Before Jxls processes the area updateCellData() will be invoked, passing the current CellData, target CellRef and data map.
The implementation can update the _CellData_ to set the required value for the cell.
For example the class below updates the total formula:

```
public class TotalCellUpdater implements CellDataUpdater {
    public void updateCellData(CellData cellData, CellRef targetCell, Context data) {
        if (cellData.isFormulaCell() && "SUM(E2)".equals(cellData.getFormula())){
            String resultFormula = String.format("SUM(E2:E%d)", targetCell.getRow());
            cellData.setEvaluationResult(resultFormula);
        }
    }
}
```
 
The key line is `cellData.setEvaluationResult(resultFormula)` which updates the cell data with the target formula.
This can be useful for example in case of streaming where it is not possible to use the Jxls standard formula processing functionality.
