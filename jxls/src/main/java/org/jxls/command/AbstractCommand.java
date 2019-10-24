package org.jxls.command;

import java.util.ArrayList;
import java.util.List;

import org.jxls.area.Area;
import org.jxls.transform.TransformationConfig;
import org.jxls.transform.Transformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Implements basic command methods and is a convenient base class for other commands
 *
 * @author Leonid Vysochyn
 */
public abstract class AbstractCommand implements Command {
    private Logger logger = LoggerFactory.getLogger(AbstractCommand.class);
    List<Area> areaList = new ArrayList<Area>();
    private String shiftMode;
    /**
     * Whether the image area is locked
     * Other commands will no longer execute in this area after locking
     * default true
     */
    private Boolean lockRange = true;

    @Override
    public Command addArea(Area area) {
        areaList.add(area);
        area.setParentCommand(this);
        return this;
    }

    @Override
    public void reset() {
        for (Area area : areaList) {
            area.reset();
        }
    }

    @Override
    public void setShiftMode(String mode) {
        if (mode != null) {
            if (mode.equalsIgnoreCase(Command.INNER_SHIFT_MODE) || mode.equalsIgnoreCase(Command.ADJACENT_SHIFT_MODE)) {
                shiftMode = mode;
            } else {
                logger.error("Cannot set cell shift mode to " + mode + " for command: " + getName());
            }
        }
    }

    @Override
    public String getShiftMode() {
        return shiftMode;
    }

    @Override
    public List<Area> getAreaList() {
        return areaList;
    }

    @Override
    public Boolean getLockRange() {
        return lockRange;
    }

    @Override
    public void setLockRange(String isLock) {
        if (isLock != null && !"".equals(isLock)) {
            this.lockRange = Boolean.valueOf(isLock);
        }
    }

    public void setLockRange(Boolean lockRange) {
        this.lockRange = lockRange;
    }

    protected Transformer getTransformer() {
        if (areaList.isEmpty()) {
            return null;
        }
        return areaList.get(0).getTransformer();
    }

    protected TransformationConfig getTransformationConfig() {
        return getTransformer().getTransformationConfig();
    }
}
