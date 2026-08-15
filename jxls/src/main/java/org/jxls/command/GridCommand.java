package org.jxls.command;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.beanutils2.PropertyUtils;
import org.jxls.area.Area;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.JxlsException;
import org.jxls.common.Size;

/**
 * The command implements a grid with dynamic columns and rows
 * 
 * @author Leonid Vysochyn
 */
public class GridCommand extends AbstractCommand {
    public static final String COMMAND_NAME = "grid";
    public static final String HEADER_VAR = "header";
    public static final String DATA_VAR = "cell";

    /** Name of a context variable containing a collection of headers */
    private String headers;
    /** Name of a context variable containing a collection of data objects for body */
    private String data;
    /** Comma separated list of object property names for each grid row */
    private String props;
    /** Comma separated list of format type cells, e.g. formatCells="Double:E1, Date:F1" */
    private String formatCells;
    private Map<String,String> cellStyleMap = new HashMap<>(); // for formatCells
    private List<String> rowObjectProps = new ArrayList<>(); // for props
    private Area headerArea;
    private Area bodyArea;

    public GridCommand() {
    }

    public GridCommand(String headers, String data) {
        this.headers = headers;
        this.data = data;
    }

    public GridCommand(String headers, String data, String props, Area headerArea, Area bodyArea) {
        this(headers, data, headerArea, bodyArea);
        this.props = props;
    }

    public GridCommand(String headers, String data, Area headerArea, Area bodyArea) {
        this.headers = headers;
        this.data = data;
        this.headerArea = headerArea;
        this.bodyArea = bodyArea;
        addArea(headerArea);
        addArea(bodyArea);
    }

    @Override
    public String getName() {
        return COMMAND_NAME;
    }

    public String getHeaders() {
        return headers;
    }

    public void setHeaders(String headers) {
        this.headers = headers;
    }

    public String getData() {
        return data;
    }

    public void setData(String data) {
        this.data = data;
    }

    public String getProps() {
        return props;
    }

    public void setProps(String props) {
        this.props = props;
        if (props != null) {
            rowObjectProps = Arrays.asList(props.replaceAll("\\s+", "").split(",")); // Remove whitespace and split into List.
        }
    }

    public String getFormatCells() {
        return formatCells;
    }

    /**
     * @param formatCells Comma-separated list of format type cells, e.g. formatCells="Double:E1, Date:F1"
     */
    public void setFormatCells(String formatCells) {
        this.formatCells = formatCells;
        cellStyleMap = new HashMap<>();
        if (formatCells == null) {
            return;
        }
        List<String> cellStyleList = Arrays.asList(formatCells.split(","));
        for (String cellStyleString : cellStyleList) {
            String[] styleCell = cellStyleString.split(":");
            cellStyleMap.put(styleCell[0].trim(), styleCell[1].trim());
        }
    }

    @Override
    public Command addArea(Area area) {
        if (areaList.size() >= 2) {
            throw new JxlsException("Cannot add any more areas to GridCommand. You can add only 1 area as a 'header' and 1 area as a 'body'.");
        }
        if (areaList.isEmpty()) {
            headerArea = area;
        } else {
            bodyArea = area;
        }
        return super.addArea(area);
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        Size headerAreaSize = processHeaders(cellRef, context);
        CellRef bodyCellRef = new CellRef(cellRef.getSheetName(), cellRef.getRow() + headerAreaSize.getHeight(), cellRef.getCol());
        Size bodyAreaSize = processBody(bodyCellRef, context);
        int gridHeight = headerAreaSize.getHeight() + bodyAreaSize.getHeight();
        int gridWidth = Math.max(headerAreaSize.getWidth(), bodyAreaSize.getWidth());
        return new Size(gridWidth, gridHeight);
    }

    private Size processHeaders(CellRef cellRef, Context context) {
        if (headerArea == null || headers == null) {
            return Size.ZERO_SIZE;
        }
        Iterable<?> headers = transformToIterableObject(this.headers, context);
        CellRef currentCell = cellRef;
        int width = 0;
        int height = 0;
        try (RunVar runVar = new RunVar(HEADER_VAR, context)) {
            for (Object header : headers) {
                runVar.put(header);
                Size size = headerArea.applyAt(currentCell, context);
                currentCell = new CellRef(currentCell.getSheetName(), currentCell.getRow(), currentCell.getCol() + size.getWidth());
                width += size.getWidth();
                height = Math.max(height, size.getHeight());
            }
        }
        return new Size(width, height);
    }

    private Size processBody(final CellRef cellRef, Context context) {
        if (bodyArea == null || data == null) {
            return Size.ZERO_SIZE;
        }
        Iterable<?> dataCollection = transformToIterableObject(data, context);

        GridCommandContext gcc = new GridCommandContext();
        gcc.currentCell = cellRef;
        boolean oldIgnoreSourceCellStyle = context.isIgnoreSourceCellStyle();
        context.setIgnoreSourceCellStyle(true);
        Map<String, String> oldCellStyleMap = context.getCellStyleMap();
        context.setCellStyleMap(this.cellStyleMap);
        try (RunVar runVar = new RunVar(DATA_VAR, context)) {
            for (Object rowObject : dataCollection) {
                Iterable<?> cellCollection;
                if (rowObject instanceof Iterable) {
                    cellCollection = (Iterable<?>) rowObject;
                } else if (rowObject instanceof Object[]) {
                    cellCollection = Arrays.asList((Object[]) rowObject);
                } else {
                    cellCollection = objectsToIterable(rowObject);
                }
                processCellCollection(cellCollection, runVar, cellRef, gcc, context);
            }
        }
        context.setIgnoreSourceCellStyle(oldIgnoreSourceCellStyle);
        context.setCellStyleMap(oldCellStyleMap);
        return new Size(gcc.totalWidth, gcc.totalHeight);
    }

    private Iterable<?> objectsToIterable(Object rowObject) {
        if (rowObjectProps.isEmpty()) {
            throw new JxlsException("Got a non-collection object type for a Grid row but object properties list is empty");
        }
        return rowObjectProps.stream().map(prop -> {
            try {
                return PropertyUtils.getProperty(rowObject, prop);
            } catch (Exception e) {
                throw new JxlsException("Failed to evaluate property " + prop + " of row object of class " + rowObject.getClass().getName(), e);
            }
        }).toList();
    }

    private void processCellCollection(Iterable<?> cellCollection, RunVar runVar, final CellRef cellRef, GridCommandContext ctx, Context context) {
        int width = 0;
        int height = 0;
        for (Object value : cellCollection) {
            runVar.put(value);
            Size size = bodyArea.applyAt(ctx.currentCell, context);
            ctx.currentCell = new CellRef(ctx.currentCell.getSheetName(), ctx.currentCell.getRow(), ctx.currentCell.getCol() + size.getWidth());
            width += size.getWidth();
            height = Math.max(height, size.getHeight());
        }
        ctx.totalWidth = Math.max(width, ctx.totalWidth);
        ctx.totalHeight = ctx.totalHeight + height;
        ctx.currentCell = new CellRef(cellRef.getSheetName(), ctx.currentCell.getRow() + height, cellRef.getCol());
    }

    static class GridCommandContext {
        CellRef currentCell;
        int totalWidth = 0;
        int totalHeight = 0;
    }
}
