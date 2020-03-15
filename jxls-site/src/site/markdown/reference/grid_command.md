Grid-Command
==============

Introduction
------------

*Grid-Command* is useful to generate a dynamic grid with a header and data row area.

Headers are passed as a collection of strings and data rows are passed as a collection of Objects or lists.

Command usage
------------------

*Grid-Command* has the following attributes

* `headers` - name of a context variable containing a collection of headers

* `data` - name of a context variable containing a collection of data 

* `props` - comma separated list of object properties for each grid row (required only if each grid row is an Object) 

* `formatCells` - comma-separated list of type-format map cells e.g. formatCells="Double:E1, Date:F1"

* `headerArea` - source xls area for headers

* `bodyArea` - source xls area for body

* `lastCell` is a common attribute for any command pointing to the last cell of the command area


The `data` variable can be of the following types

1. `Collection<Collection<Object>>` - here each inner collection contains cell values for the corresponding row
2. `Collection<Object>` - here each collection item is an object containing the data for the corresponding row. 
In this case you have to specify `props` attribute to define which object properties should be used to set the data for a particular cell.

When iterating by a collection of headers *Grid-Command* puts each header into the context under `header` key.
During iteration by data rows each cell item is put into the context under `cell` key.

So in Excel template *Grid-Command* requires only 2 cells - one is for header cell and one is for data row cell. 
The header cell can be defined as
     
     ${header}

The data row cell can be defined as

    ${cell}

See an example of *Grid-Command* usage in [Dynamic Grid Sample](../samples/dynamic_grid.html)

Transformer Support Note
--------------------------
Please note that *Grid-Command* is currently supported only in POI transformer so you have to use POI if you are using *Grid-Command*.
