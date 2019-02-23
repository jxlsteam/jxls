package org.jxls.command;

import org.jxls.area.Area;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class UpdateCellCommand extends AbstractCommand {
    public static final String COMMAND_NAME = "updateCell";
    private static Logger logger = LoggerFactory.getLogger(UpdateCellCommand.class);

    private Area area;
    private CellDataUpdater cellDataUpdater;
    private String updater;


    @Override
    public String getName() {
        return COMMAND_NAME;
    }

    public String getUpdater() {
        return updater;
    }

    public void setUpdater(String updater) {
        this.updater = updater;
    }

    public CellDataUpdater getCellDataUpdater() {
        return cellDataUpdater;
    }

    public void setCellDataUpdater(CellDataUpdater cellDataUpdater) {
        this.cellDataUpdater = cellDataUpdater;
    }

    @Override
    public Command addArea(Area area) {
        String message = "You can add only a single-cell area to '" + COMMAND_NAME + "' command";
        if (areaList.size() >= 1) {
            throw new IllegalArgumentException(message);
        }
        if (area != null && area.getSize().getHeight() != 1 && area.getSize().getWidth() != 1) {
            throw new IllegalArgumentException(message);
        }
        this.area = area;
        return super.addArea(area);
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        CellDataUpdater cellDataUpdater = createCellDataUpdater(context);
        CellRef srcCell = area.getStartCellRef();
        CellData cellData = area.getTransformer().getCellData(srcCell);
        if( cellDataUpdater != null ) {
            cellDataUpdater.updateCellData(cellData, cellRef, context);
        }
        return area.applyAt(cellRef, context);
    }

    private CellDataUpdater createCellDataUpdater(Context context) {
        if( updater==null ){
            logger.warn("updater attribute is not set");
            return null;
        }
        if (context.getVar(updater) instanceof CellDataUpdater) {
            return (CellDataUpdater) context.getVar(updater);
        } else {
            logger.warn("CellDataUpdater is null for updater ='{}'", updater);
            return null;
        }
    }
}
