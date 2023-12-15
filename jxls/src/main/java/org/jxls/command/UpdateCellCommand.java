package org.jxls.command;

import org.jxls.area.Area;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.JxlsException;
import org.jxls.common.Size;

/**
 * Allows to update a cell data dynamically for any cell
 */
public class UpdateCellCommand extends AbstractCommand {
    public static final String COMMAND_NAME = "updateCell";
    private Area area;
    private String updater;
    private CellDataUpdater cellDataUpdater;

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
        String message = "You can only add a single cell area to '" + COMMAND_NAME + "' command!";
        if (areaList.size() >= 1) {
            throw new JxlsException(message);
        }
        if (area != null && area.getSize().getHeight() != 1 && area.getSize().getWidth() != 1) {
            throw new JxlsException(message);
        }
        this.area = area;
        return super.addArea(area);
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        CellDataUpdater cellDataUpdater = createCellDataUpdater(context);
        if (cellDataUpdater != null) {
            CellRef srcCell = area.getStartCellRef();
            CellData cellData = area.getTransformer().getCellData(srcCell);
            cellDataUpdater.updateCellData(cellData, cellRef, context);
        }
        return area.applyAt(cellRef, context);
    }

    private CellDataUpdater createCellDataUpdater(Context context) {
        if (updater == null) {
        	throw new JxlsException("updater must not be null");
        }
        Object updaterInstance = context.getVar(updater);
        if (updaterInstance instanceof CellDataUpdater u) {
            return u;
        } else {
        	throw new JxlsException("Value of var '" + updater + "' must implement CellDataUpdater!");
        }
    }
}
