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
    
    public IfCommand(String condition, Pos pos, Size initialSize, Command ifArea, Command elseArea){
        super(pos, initialSize);
        this.condition = condition;
        this.ifArea = ifArea != null ? ifArea : BaseComand.EMPTY_COMAND;
        this.elseArea = elseArea != null ? elseArea : BaseComand.EMPTY_COMAND;
    }

    public IfCommand(String condition, Pos pos, Size initialSize,  BaseComand ifArea) {
        this(condition, pos, initialSize, ifArea, BaseComand.EMPTY_COMAND);
    }

    public Size applyAt(Pos pos, Context context) {
        conditionResult = isConditionTrue(context);
        if( conditionResult ){
            return ifArea.applyAt(pos, context);
        }else{
            return elseArea.applyAt(pos, context);
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
