package org.jxls.command;

import org.jxls.area.Area;
import org.jxls.area.XlsArea;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.util.Util;

/**
 * Implements if-else logic
 * 
 * @author Leonid Vysochyn
 */
public class IfCommand extends AbstractCommand {
    public static final String COMMAND_NAME = "if";
    private String condition;
    private Area ifArea = XlsArea.EMPTY_AREA;
    private Area elseArea = XlsArea.EMPTY_AREA;

    public IfCommand() {
    }

    /**
     * @param condition JEXL expression for boolean condition to evaluate
     */
    public IfCommand(String condition) {
        this.condition = condition;
    }

    public IfCommand(String condition, XlsArea ifArea) {
        this(condition, ifArea, XlsArea.EMPTY_AREA);
    }

    public IfCommand(String condition, Area ifArea, Area elseArea) {
        this.condition = condition;
        this.ifArea = ifArea != null ? ifArea : XlsArea.EMPTY_AREA;
        this.elseArea = elseArea != null ? elseArea : XlsArea.EMPTY_AREA;
        super.addArea(this.ifArea);
        super.addArea(this.elseArea);
    }

    @Override
    public String getName() {
        return COMMAND_NAME;
    }

    /**
     * Gets test condition as JEXL expression string
     * @return test condition
     */
    public String getCondition() {
        return condition;
    }

    /**
     * Sets test condition as JEXL expression string
     * @param condition condition to test
     */
    public void setCondition(String condition) {
        this.condition = condition;
    }

    /**
     * Gets an area to render when the condition is evaluated to 'true'
     * @return if Area
     */
    public Area getIfArea() {
        return ifArea;
    }

    /**
     * Gets an area to render when the condition is evaluated to 'false'
     * @return else Area
     */
    public Area getElseArea() {
        return elseArea;
    }

    @Override
    public Command addArea(Area area) {
        if (areaList.size() >= 2) {
            throw new IllegalArgumentException(
                    "Cannot add any more areas to this IfCommand. You can add only 1 area for 'if' part and 1 area for 'else' part");
        }
        if (areaList.isEmpty()) {
            ifArea = area;
        } else {
            elseArea = area;
        }
        return super.addArea(area);
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        Boolean conditionResult = Util.isConditionTrue(getTransformationConfig().getExpressionEvaluator(), condition, context);
        if (conditionResult.booleanValue()) {
            return ifArea.applyAt(cellRef, context);
        } else {
            return elseArea.applyAt(cellRef, context);
        }
    }
}
