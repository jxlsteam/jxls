package org.jxls.command;

import org.jxls.area.Area;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.util.Util;

import java.util.Collection;

/**
 * The command implements a grid with dynamic columns and rows
 * Created by Leonid Vysochyn on 25-Jun-15.
 */
public class GridCommand extends AbstractCommand {
    public static final String COMMAND_NAME = "grid";
    public static final String HEADER_VAR_NAME = "header";

    String items;
    String headerVar;
    String dataVar;
    Area headerArea;
    Area bodyArea;

    @Override
    public String getName() {
        return COMMAND_NAME;
    }

    public String getItems() {
        return items;
    }

    public void setItems(String items) {
        this.items = items;
    }

    public String getHeaderVar() {
        return headerVar;
    }

    public void setHeaderVar(String headerVar) {
        this.headerVar = headerVar;
    }

    public String getDataVar() {
        return dataVar;
    }

    public void setDataVar(String dataVar) {
        this.dataVar = dataVar;
    }

    @Override
    public Command addArea(Area area) {
        if( areaList.size() >= 2 ){
            throw new IllegalArgumentException("Cannot add any more areas to GridCommand. You can add only 1 area as a 'header' and 1 area as a 'body'");
        }
        if(areaList.isEmpty()){
            headerArea = area;
        }else {
            bodyArea = area;
        }
        return super.addArea(area);
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        Size headerAreaSize = processHeaders(cellRef, context);
        return headerAreaSize;
    }

    private Size processHeaders(CellRef cellRef, Context context) {

        if(headerArea == null || headerVar == null){
            return Size.ZERO_SIZE;
        }
        Collection headers = Util.transformToCollectionObject(headerVar, context);
        CellRef currentCell = cellRef;
        JexlExpressionEvaluator selectEvaluator = null;
        int width = 0;
        int height = 0;
        int index = 0;

        for( Object header : headers){
            context.putVar(HEADER_VAR_NAME, header);
            Size size = headerArea.applyAt(currentCell, context);

            currentCell = new CellRef(currentCell.getSheetName(), currentCell.getRow(), currentCell.getCol() + size.getWidth());
            width += size.getWidth();
                height = Math.max( height, size.getHeight() );
            context.removeVar(HEADER_VAR_NAME);
        }

        return new Size(width, height);
    }

}
