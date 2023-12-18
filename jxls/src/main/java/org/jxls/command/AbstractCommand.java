package org.jxls.command;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import org.jxls.area.Area;
import org.jxls.common.Context;
import org.jxls.common.JxlsException;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.logging.JxlsLogger;
import org.jxls.transform.TransformationConfig;
import org.jxls.transform.Transformer;

/**
 * Implements basic command methods and is a convenient base class for other commands
 *
 * @author Leonid Vysochyn
 */
public abstract class AbstractCommand implements Command {
    protected List<Area> areaList = new ArrayList<Area>();
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
    	areaList.forEach(area -> area.reset());
    }

    @Override
    public void setShiftMode(String mode) {
    	if (Command.INNER_SHIFT_MODE.equals(mode) || Command.ADJACENT_SHIFT_MODE.equals(mode)) {
    		shiftMode = mode;
    	} else {
			throw new IllegalArgumentException("Use \"" + Command.INNER_SHIFT_MODE + "\" or \""
					+ Command.ADJACENT_SHIFT_MODE + "\" for shiftMode.");
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
        if (isLock != null && !isLock.isEmpty()) {
            this.lockRange = Boolean.valueOf(isLock);
        }
    }

    public void setLockRange(Boolean lockRange) {
        this.lockRange = lockRange;
    }

    protected Transformer getTransformer() {
        return areaList.isEmpty() ? null : areaList.get(0).getTransformer();
    }
    
    protected TransformationConfig getTransformationConfig() {
        return getTransformer().getTransformationConfig();
    }

    protected JxlsLogger getLogger() {
        Transformer transformer = getTransformer();
        if (transformer == null) {
            throw new JxlsException("Command has no transformer and can not write to log");
        }
        return transformer.getLogger();
    }

    /**
     * Evaluates the passed collection name into an {@link Iterable} object
     * @param expressionEvaluator -
     * @param collectionName -
     * @param context -
     * @return an iterable object from the {@link Context} under given name
     */
    protected Iterable<Object> transformToIterableObject(String collectionName, Context context) {
        Object collectionObject = getExpressionEvaluator().evaluate(collectionName, context.toMap());
        if (collectionObject == null) {
            return Collections.emptyList();
        } else if (collectionObject instanceof Object[]) {
            return Arrays.asList((Object[])/*cast is important*/ collectionObject);
        } else if (collectionObject instanceof Iterable) {
            @SuppressWarnings("unchecked")
            Iterable<Object> iterable = (Iterable<Object>) collectionObject;
            return iterable;
        }
        throw new JxlsException(collectionName + " expression is not a collection or an array");
    }

    protected ExpressionEvaluator getExpressionEvaluator() {
        return getTransformationConfig().getExpressionEvaluator();
    }
}
