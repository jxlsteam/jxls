Area listener
=============

Introduction
------------

Area listeners can be used to perform an additional area processing in response to area transformation events.
For example you may want to highlight some rows or cells depending on the data.

AreaListener interface
----------------------

*AreaListener* interface looks like this

    public interface AreaListener {
        void beforeApplyAtCell(CellRef cellRef, Context context);
        void afterApplyAtCell(CellRef cellRef, Context context);
        void beforeTransformCell(CellRef srcCell, CellRef targetCell, Context context);
        void afterTransformCell(CellRef srcCell, CellRef targetCell, Context context);
    }


When a cell in the corresponding xls area is being transformed a corresponding method is being invoked.
Each listener method gets a cell reference to the cell being transformed and the context. Transformation listener methods also get the target cell reference.

* `beforeApplyAtCell()` method is invoked just before the cell processing is going start
* `afterApplyAtCell()` method is invoked after the cell processing already completed
* `beforeTransformCell()` method is invoked when the cell is about to be transformed by a transformer
* `afterTransformCell()` method is invoked after the cell was transformed by a transformer

See [Area listener example](../samples/area_listener.html) to see it in action.


