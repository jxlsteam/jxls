package com.jxls.writer.command;

import com.jxls.writer.*;
import com.jxls.writer.expression.ExpressionEvaluator;
import com.jxls.writer.expression.JexlExpressionEvaluator;

/**
 * Date: Sep 11, 2009
 *
 * @author Leonid Vysochyn
 */
public class IfCommand extends AbstractCommand {

    String condition;
    Boolean conditionResult;
    Command ifArea;
    Command elseArea;
    
    public IfCommand(String condition, Size initialSize, Area ifArea, Area elseArea){
        super(initialSize);
        this.condition = condition;
        this.ifArea = ifArea != null ? ifArea : BaseArea.EMPTY_AREA;
        this.elseArea = elseArea != null ? elseArea : BaseArea.EMPTY_AREA;
    }

    public IfCommand(String condition, Size initialSize,  BaseArea ifArea) {
        this(condition, initialSize, ifArea, BaseArea.EMPTY_AREA);
    }

    public Size applyAt(Cell cell, Context context) {
        conditionResult = isConditionTrue(context);
        if( conditionResult ){
            return ifArea.applyAt(cell, context);
        }else{
            return elseArea.applyAt(cell, context);
        }
    }

    public Boolean isConditionTrue(Context context){
        ExpressionEvaluator expressionEvaluator = new JexlExpressionEvaluator(context.toMap());
        Object conditionResult = expressionEvaluator.evaluate(condition);
        if( !(conditionResult instanceof Boolean) ){
            throw new RuntimeException("If command condition is not a boolean value - " + condition);
        }
        return (Boolean)conditionResult;
    }
}
