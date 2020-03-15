Main Concepts
=============

Jxls is based on the following main concepts

* XlsArea

* Command

* Transformer

Let's discuss each of these concepts in detail.

XlsArea
-------

*XlsArea* represents a rectangular area in an Excel file. It may be defined using a cell range or by specifying a start cell and
a size (number of columns and rows) of the area. *XlsArea* includes all the Excel cells in the specified range.

Each *XlsArea* may have a set of *Commands* associated with it which will be executed during the area processing by Jxls engine.
*XlsArea* may have child areas which are nested in it. Each child area is also *XlsArea* with its own *Commands* and may have its own child areas.

*XlsArea* can be defined using the following ways

* By using a special markup syntax in an Excel template. Jxls provides one default mark-up with its *XlsCommentAreaBuilder*. Custom mark-ups may be defined if needed.
* By using XML configuration. Jxls provides *XmlAreaBuilder* class as a default implementation of XML mark-up.
* By using Jxls Java API

You can find more details and usage examples in [here](xls_area.html)

Command
-------
*Command* represents a transformation action on a single or multiple *XlsAreas*.

The corresponding Java interface looks like this

    public interface Command {
        String getName();
        List<Area> getAreaList();
        Command addArea(Area area);
        Size applyAt(CellRef cellRef, Context context);
        void reset();
    }

The main method of any *Command* is `Size applyAt(CellRef cellRef, Context context)` . The method applies the *Command* action
at cell `cellRef` passing the data with `context` variable. *Context* works like a map and is used to pass the data to a command.
The method returns the new dimensions of the transformed area as a `Size` object.

Currently Jxls provides the following built-in commands

* [Each-Command](each_command.html) - iterates over a collection (Iterable<?> or array) of items and processes the command area multiple times putting a corresponding item into the context
* [If-Command](if_command.html) - processes the command areas based on a condition
* [Image-Command](image_command.html) - renders an image
* [MergeCells-Command](merge_cells_command.html) - merges cells


You can also define custom commands as explained in [Custom-Command](custom_command.html)

The passing data to a *Command* is performed through the *Context* object which works as a map with keys referred in XLS template and values set to the data.

Transformer
-----------
*Transformer* interface allows *XlsArea* to interact with Excel independently of any specific implementation.
It means that by providing different implementations of *Transformer* interface we can use different underlying Java-to-Excel libraries.

The interface looks like this

    public interface Transformer {
        void transform(CellRef srcCellRef, CellRef targetCellRef, Context context, boolean updateRowHeight);

        void setFormula(CellRef cellRef, String formulaString);

        Set<CellData> getFormulaCells();

        CellData getCellData(CellRef cellRef);

        List<CellRef> getTargetCellRef(CellRef cellRef);

        void resetTargetCellRefs();

        void resetArea(AreaRef areaRef);

        void clearCell(CellRef cellRef);

        List<CellData> getCommentedCells();

        void addImage(AreaRef areaRef, byte[] imageBytes, ImageType imageType);

        void write() throws IOException;

        TransformationConfig getTransformationConfig();

        void setTransformationConfig(TransformationConfig transformationConfig);

        boolean deleteSheet(String sheetName);

        void setHidden(String sheetName, boolean hidden);

        void updateRowHeight(String srcSheetName, int srcRowNum, String targetSheetName, int targetRowNum);
    }

Although it looks like a lot of methods but many of them are already implemented in the base abstract class *AbstractTransformer*
and can be just inherited if one needs to support a new java-to-excel implementation.

Jxls provides a built-in implementation of the Transformer interface *PoiTransformer* which is based on the well-known [Apache POI](https://poi.apache.org/) library.

