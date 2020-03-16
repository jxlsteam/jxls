*UpdateCell-Command*
===================

Introduction
------------

*UpdateCell-Command* allows you to customize the processing for a particular cell.

Command Attributes
------------------

The *UpdateCell-Command*  has the following attributes

* `updater` is a name of the key in the context containing *CellDataUpdater* implementation 

* `lastCell` is a common attribute for any command pointing to the last cell of the command area

The _CellDataUpdater_ interface looks like this

    public interface CellDataUpdater {
        void updateCellData(CellData cellData, CellRef targetCell, Context context);
    }
    
Before transforming the area *UpdateCell-Command* will invoke the _updateCellData_ method passing the current _CellData_, the target cell and the current context.

The implementation can update the _CellData_ to set the required value for the cell.

For example the class below updates the total formula

    class TotalCellUpdater implements CellDataUpdater{
        public void updateCellData(CellData cellData, CellRef targetCell, Context context) {
            if( cellData.isFormulaCell() && cellData.getFormula().equals("SUM(E2)")){
                String resultFormula = String.format("SUM(E2:E%d)", targetCell.getRow());
                cellData.setEvaluationResult(resultFormula);
            }
        }
    }
 
 The key line is *cellData.setEvaluationResult(resultFormula)* which updates the cell data with the target formula.
 
 This can be useful for example in SXSSF processing where it is not possible to use the standard Jxls formula processing functionality.
  
 See *SxssfDemo.java* as an Update Cell command example in org.jxls.examples.
 

Excel markup
-------------

To create *UpdateCell Command* in Excel template create a cell comment like this

    jx:updateCell(lastCell="E4"  updater="totalCellUpdater")

The `lastCell` attribute defines the last cell of the command area.

The `updater` attribute is set to _totalCellUpdater_. The _totalCellUpdater_ must be put into the context before the processing 

        Context context = new Context();
        context.putVar("totalCellUpdater", new TotalCellUpdater());


