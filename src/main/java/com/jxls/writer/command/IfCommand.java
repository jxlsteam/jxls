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
    
    public IfCommand(String condition, Cell cell, Size initialSize, Command ifArea, Command elseArea){
        super(cell, initialSize);
        this.condition = condition;
        this.ifArea = ifArea != null ? ifArea : BaseCommand.EMPTY_COMMAND;
        this.elseArea = elseArea != null ? elseArea : BaseCommand.EMPTY_COMMAND;
    }

    public IfCommand(String condition, Cell cell, Size initialSize,  BaseCommand ifArea) {
        this(condition, cell, initialSize, ifArea, BaseCommand.EMPTY_COMMAND);
    }

    public Size applyAt(Cell cell, Context context) {
        conditionResult = isConditionTrue(context);
        if( conditionResult ){
            return ifArea.applyAt(cell, context);
        }else{
            return elseArea.applyAt(cell, context);
        }
    }

    public Size getSize(Context context) {
        Boolean conditionResult = isConditionTrue(context);
        return conditionResult ? ifArea.getSize(context) : elseArea.getSize(context);
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
